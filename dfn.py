"""
dfn.py - DocxFormulaNumbering

Word 文档公式自动编号工具

功能:
- 自动识别 Word 文档中的单行公式
- 添加带章节号的编号格式 (1-1), (1-2), (2-1)...
- 支持简单连续编号格式 (1), (2), (3)...
- 创建书签用于交叉引用
- 支持 eqArr 格式（美观的公式编号显示）

用法:
    python dfn.py <输入 docx> [输出 docx] [--simple] [--t1]

示例:
    # 默认模式：章节号 - 序号格式
    python dfn.py thesis.docx output.docx
    
    # 简单模式：连续编号
    python dfn.py thesis.docx output.docx --simple
    
    # eqArr 格式（需要 t1.docx 模板）
    python dfn.py thesis.docx output.docx --t1
    
    # 覆盖原文件
    python dfn.py document.docx

作者：FaustSherpad
版本：1.0.0
"""

import win32com.client
import re
from pathlib import Path
import argparse
import zipfile
import shutil
import tempfile
import os


# ==================== 工具函数 ====================

def is_display_formula(para) -> bool:
    """判断是否为单行公式（显示公式）"""
    if para.Range.OMaths.Count == 0:
        return False

    para_xml = para.Range.XML
    texts = []
    for t in re.findall(r'<w:t[^>]*>(.*?)</w:t>', para_xml, re.DOTALL):
        if t.strip():
            texts.append(t.strip())

    visible_texts = [t for t in texts if '<m:' not in t and '</m:' not in t]
    non_formula_text_len = sum(len(t) for t in visible_texts)
    has_label = any(':' in t for t in visible_texts)
    is_inline = non_formula_text_len > 20 or has_label

    return not is_inline


def get_chapter_number(para, word) -> int:
    """使用 Word COM 接口获取当前段落所在章节的编号"""
    try:
        original_range = word.Selection.Range.Duplicate
        para.Range.Select()
        word.Selection.Collapse(1)

        temp_field = word.ActiveDocument.Fields.Add(
            Range=word.Selection.Range,
            Type=-1,
            Text='STYLEREF "标题 1" \\n',
            PreserveFormatting=False
        )

        temp_field.Update()
        chapter_text = temp_field.Result.Text.strip()
        temp_field.Delete()
        original_range.Select()

        if chapter_text and ('错误' in chapter_text or 'Error' in chapter_text):
            return 0

        if chapter_text:
            try:
                return int(chapter_text)
            except ValueError:
                return 0
        return 0

    except:
        return 0


def check_multilevel_list_configured(doc, word) -> bool:
    """检查文档是否正确配置了多级列表"""
    try:
        original_range = word.Selection.Range.Duplicate
        doc.Range(0, 0).Select()

        temp_field = word.ActiveDocument.Fields.Add(
            Range=word.Selection.Range,
            Type=-1,
            Text='STYLEREF "标题 1" \\n',
            PreserveFormatting=False
        )

        temp_field.Update()
        chapter_text = temp_field.Result.Text.strip()
        temp_field.Delete()
        original_range.Select()

        if chapter_text:
            if '错误' in chapter_text or 'Error' in chapter_text:
                pass
            else:
                try:
                    if int(chapter_text) > 0:
                        return True
                except ValueError:
                    pass

        for para in doc.Paragraphs:
            try:
                style_name = para.Style.NameLocal
                if "标题 1" in style_name or "Heading 1" in style_name:
                    if para.Range.ListFormat.ListValue > 0:
                        return True
            except:
                pass

        return False

    except Exception as e:
        print(f'检查多级列表配置失败：{e}')
        return False


# ==================== eqArr 模板处理 ====================

def extract_eqarr_template(template_path: str) -> dict:
    """
    从 t1.docx 模板中提取 eqArr 结构
    
    返回:
        dict: 包含模板各部分的字典
    """
    with zipfile.ZipFile(template_path, 'r') as zip_ref:
        template_xml = zip_ref.read('word/document.xml').decode('utf-8')
    
    # 提取 oMathPara 模板
    omath_para_match = re.search(
        r'(<m:oMathPara>.*?</m:oMathPara>)',
        template_xml,
        re.DOTALL
    )
    
    if not omath_para_match:
        raise ValueError("模板中未找到 oMathPara 元素")
    
    omath_template = omath_para_match.group(1)
    
    # 提取公式内容（# 前的部分）
    formula_match = re.search(r'<m:t[^>]*>([^#]*)#', omath_template, re.DOTALL)
    if not formula_match:
        raise ValueError("模板中未找到公式内容")
    
    original_formula = formula_match.group(1).strip()
    
    # 提取编号
    number_match = re.search(r'<m:t>(\d+)</m:t>', omath_template)
    if not number_match:
        raise ValueError("模板中未找到编号")
    
    original_number = number_match.group(1)
    
    # 提取 eqArr 环境的完整结构
    eqarr_match = re.search(r'(<m:eqArr>.*?</m:eqArr>)', omath_template, re.DOTALL)
    if not eqarr_match:
        raise ValueError("模板中未找到 eqArr 元素")
    
    eqarr_template = eqarr_match.group(1)
    
    return {
        'omath_para': omath_template,
        'eqarr': eqarr_template,
        'original_formula': original_formula,
        'original_number': original_number,
    }


def create_eqarr_formula(template: dict, formula_text: str, number: str) -> str:
    """
    使用模板创建 eqArr 格式的公式
    
    参数:
        template: 从 extract_eqarr_template 返回的模板字典
        formula_text: 公式内容
        number: 编号
    
    返回:
        str: eqArr 格式的 oMathPara XML
    """
    # 从 eqarr 模板开始
    eqarr_xml = template['eqarr']
    
    # 替换公式内容
    eqarr_xml = eqarr_xml.replace(
        f'>{template["original_formula"]}#<',
        f'>{formula_text}#<'
    )
    
    # 替换编号（只替换第一个匹配的）
    eqarr_xml = re.sub(
        r'<m:t>\d+</m:t>',
        f'<m:t>{number}</m:t>',
        eqarr_xml,
        count=1
    )
    
    # 构建完整的 oMathPara
    omath_para = f'<m:oMathPara><m:oMath>{eqarr_xml}</m:oMath></m:oMathPara>'
    
    return omath_para


def extract_formula_text(omath_content: str) -> str:
    """
    从 oMath 内容中提取纯文本公式
    
    将所有 <m:t> 元素的内容连接起来，直到遇到 # 分隔符
    """
    # 查找所有 m:t 元素
    texts = re.findall(r'<m:t[^>]*>([^<]*)</m:t>', omath_content)
    
    # 连接文本，直到遇到 #
    result = []
    for t in texts:
        if '#' in t:
            # 找到#，提取#前的内容
            idx = t.index('#')
            result.append(t[:idx])
            break
        else:
            result.append(t)
    
    return ''.join(result).strip()


def extract_formula_number(omath_content: str, simple_mode: bool = False) -> str:
    """
    从 oMath 内容中提取公式编号
    
    参数:
        omath_content: oMath 内容 XML
        simple_mode: 简单模式（无章节号）
    
    返回:
        str: 编号（简单模式返回纯数字，默认模式返回"章节 - 序号"格式）
    """
    chapter = None
    num = None
    
    # 查找书签名：eq{chap}_{num} 或 eq{num}
    bookmark_match = re.search(r'w:name="eq(\d+)_(\d+)"', omath_content)
    if bookmark_match:
        chapter = bookmark_match.group(1)
        num = bookmark_match.group(2)
    else:
        # 尝试简单模式书签：eq{num}
        bookmark_simple_match = re.search(r'w:name="eq(\d+)"', omath_content)
        if bookmark_simple_match:
            num = bookmark_simple_match.group(1)
            chapter = None
    
    # 如果书签未找到，尝试从 SEQ 域提取
    if num is None:
        # 查找 SEQ 域中的章节号：SEQ 式 \* ARABIC \s N
        seq_chapter_match = re.search(r'SEQ\s+式\s+\\*\s*ARABIC\s+\\s\s+(\d+)', omath_content)
        if seq_chapter_match:
            chapter = seq_chapter_match.group(1)
        
        # 查找域结果中的序号
        # 在 SEQ 域之后查找第一个数字
        seq_match = re.search(r'SEQ\s+式\s+\\*\s*ARABIC', omath_content)
        if seq_match:
            seq_end = seq_match.end()
            after_seq = omath_content[seq_end:seq_end+200]
            num_match = re.search(r'<w:t[^>]*>(\d+)</w:t>', after_seq)
            if num_match:
                num = num_match.group(1)
    
    # 如果仍然未找到，返回 None
    if num is None:
        return None
    
    # 根据模式返回编号
    if simple_mode or chapter is None:
        return num
    else:
        return f"{chapter}-{num}"


def convert_formula_in_xml(xml_content: str, template: dict, simple_mode: bool = False) -> str:
    """
    将 XML 中的公式从标准格式转换为 eqArr 格式

    标准格式：<m:r><m:t>公式#</m:t></m:r>...域...<m:r><m:t>)</m:t></m:r>
    eqArr 格式：<m:eqArr>...</m:eqArr>
    
    参数:
        xml_content: XML 内容
        template: 模板字典
        simple_mode: 简单模式（无章节号）
    """
    
    # 查找所有 oMathPara 元素
    omath_paras = re.findall(r'(<m:oMathPara>.*?</m:oMathPara>)', xml_content, re.DOTALL)
    
    print(f"找到 {len(omath_paras)} 个 oMathPara 元素")
    
    for i, para in enumerate(omath_paras):
        # 检查是否已经包含 eqArr
        if '<m:eqArr>' in para:
            print(f"  oMathPara {i}: 已包含 eqArr，跳过")
            continue
        
        # 提取公式内容
        formula_text = extract_formula_text(para)
        if not formula_text:
            print(f"  oMathPara {i}: 未找到公式内容，跳过")
            continue
        
        # 提取编号
        number = extract_formula_number(para, simple_mode)
        if not number:
            print(f"  oMathPara {i}: 未找到编号，跳过")
            continue
        
        print(f"  oMathPara {i}: 公式='{formula_text}', 编号={number}")
        
        # 创建 eqArr 格式
        eqarr_xml = create_eqarr_formula(template, formula_text, number)
        
        # 替换原 XML 中的 oMathPara
        xml_content = xml_content.replace(para, eqarr_xml, 1)
    
    return xml_content


def postprocess_eqarr_format(docx_path: str, template_path: str, output_path: str = None, simple_mode: bool = False) -> bool:
    """
    后处理：将 docx 文件中的公式转换为 eqArr 格式

    参数:
        docx_path: 已编号的 docx 文件路径
        template_path: t1.docx 模板路径
        output_path: 输出路径（如果为 None，则覆盖原文件）
        simple_mode: 简单模式（无章节号）

    返回:
        bool: 是否成功
    """
    if output_path is None:
        output_path = docx_path

    try:
        # 提取模板
        template = extract_eqarr_template(template_path)
        print(f"模板提取成功：公式='{template['original_formula']}', 编号='{template['original_number']}'")

        # 读取文档 XML
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            xml_content = zip_ref.read('word/document.xml').decode('utf-8')

        # 转换公式
        print("正在转换公式格式...")
        new_xml = convert_formula_in_xml(xml_content, template, simple_mode)

        # 检查是否有转换
        if new_xml == xml_content:
            print("警告：未检测到需要转换的公式")
        else:
            print("公式格式转换完成")

        # 写入输出文件
        with zipfile.ZipFile(output_path, 'w') as zip_out:
            for item in zipfile.ZipFile(docx_path, 'r').infolist():
                if item.filename == 'word/document.xml':
                    zip_out.writestr(item, new_xml)
                else:
                    zip_out.writestr(item, zipfile.ZipFile(docx_path, 'r').read(item.filename))

        # 验证
        with zipfile.ZipFile(output_path, 'r') as zip_ref:
            new_xml_check = zip_ref.read('word/document.xml').decode('utf-8')
            has_eqarr = '<m:eqArr>' in new_xml_check
            has_delim = '<m:d>' in new_xml_check
            print(f"验证结果：eqArr={has_eqarr}, m:d={has_delim}")

        return True
        
    except Exception as e:
        print(f"后处理失败：{e}")
        import traceback
        traceback.print_exc()
        return False


# ==================== 主处理函数 ====================

def add_formula_numbering(doc_path: str, output_path: str = None, simple_mode: bool = False, use_eqarr: bool = False):
    """
    为 Word 文档中的单行公式添加编号和书签

    参数:
        doc_path: 输入文档路径
        output_path: 输出文档路径
        simple_mode: 简单模式（连续编号，不含章节号）
        use_eqarr: 使用 eqArr 格式（t1 格式）
    """
    if output_path is None:
        output_path = doc_path

    word = None
    doc = None
    temp_file = None
    
    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(str(doc_path))

        # 如果不是简单模式，检查文档是否配置了多级列表
        if not simple_mode:
            print('检查文档多级列表配置...')
            if not check_multilevel_list_configured(doc, word):
                doc.Close(SaveChanges=False)
                raise ValueError(
                    '文档未配置多级列表！\n\n'
                    '解决方法:\n'
                    '1. 为章节标题应用 "标题 1" 样式\n'
                    '2. 使用 "开始" → "段落" → "多级列表" 应用编号格式\n'
                    '3. 重新运行此脚本\n\n'
                    '或者使用 --simple 参数跳过章节号检查'
                )
            print('多级列表配置检查通过')

        # 统计章节和公式
        chapter = 0
        formula_paras = []
        formula_count = 0
        global_formula_count = 0

        for i, para in enumerate(doc.Paragraphs):
            para_text = para.Range.Text.strip()

            is_heading = False
            try:
                style_name = para.Style.NameLocal
                if "标题 1" in style_name or "Heading 1" in style_name:
                    is_heading = True
            except:
                pass

            if is_heading:
                chapter += 1
                formula_count = 0
                continue

            if para.Range.OMaths.Count > 0:
                if is_display_formula(para):
                    para_xml = para.Range.XML
                    if 'SEQ' not in para_xml and 'STYLEREF' not in para_xml:
                        formula_count += 1
                        global_formula_count += 1

                        if simple_mode:
                            formula_paras.append((i, 0, global_formula_count, global_formula_count))
                        else:
                            formula_paras.append((i, chapter, formula_count, global_formula_count))

        if chapter == 0 and not simple_mode:
            doc.Close(SaveChanges=False)
            raise ValueError(
                '文档未配置多级列表！\n\n'
                '解决方法:\n'
                '1. 为章节标题应用 "标题 1" 样式\n'
                '2. 使用 "开始" → "段落" → "多级列表" 应用编号格式\n'
                '3. 重新运行此脚本\n\n'
                '或者使用 --simple 参数跳过章节号检查'
            )

        print(f'\n找到 {len(formula_paras)} 个单行公式')

        # 从后往前处理
        processed_count = 0
        for para_idx, chap, num, global_num in reversed(formula_paras):
            para = doc.Paragraphs(para_idx + 1)

            try:
                formula = para.Range.OMaths.Item(1)
                formula.Range.Select()
                word.Selection.Collapse(0)

                # 1. 添加 # 分隔符
                word.Selection.TypeText('#')
                word.Selection.Collapse(0)

                # 2. 添加左括号
                word.Selection.TypeText('(')
                start_pos = word.Selection.End
                word.Selection.Collapse(0)

                if simple_mode:
                    field_code = 'SEQ 式 \\* ARABIC'
                    field = doc.Fields.Add(
                        Range=word.Selection.Range,
                        Type=-1,
                        Text=field_code,
                        PreserveFormatting=False
                    )
                    field.Select()
                    word.Selection.MoveRight(Unit=2, Count=1)
                else:
                    field1 = doc.Fields.Add(
                        Range=word.Selection.Range,
                        Type=-1,
                        Text='STYLEREF "标题 1" \\n',
                        PreserveFormatting=False
                    )
                    field1.Select()
                    word.Selection.MoveRight(Unit=2, Count=1)

                    word.Selection.TypeText('-')
                    word.Selection.Collapse(0)

                    field_code = f'SEQ 式 \\* ARABIC \\s {chap}'
                    field2 = doc.Fields.Add(
                        Range=word.Selection.Range,
                        Type=-1,
                        Text=field_code,
                        PreserveFormatting=False
                    )
                    field2.Select()
                    word.Selection.MoveRight(Unit=2, Count=1)

                # 3. 添加右括号
                end_pos = word.Selection.Start
                word.Selection.TypeText(')')

                # 4. 创建书签
                bookmark_range = doc.Range(start_pos, end_pos)
                if simple_mode:
                    bookmark_name = f'eq{global_num}'
                else:
                    bookmark_name = f'eq{chap}_{num}'
                doc.Bookmarks.Add(Name=bookmark_name, Range=bookmark_range)

                print(f'  段落 {para_idx}: 书签 {bookmark_name}')
                processed_count += 1

            except Exception as e:
                print(f'段落 {para_idx} 处理失败：{e}')

        print(f'\n处理完成，共处理 {processed_count} 个公式')

        # 更新所有域
        doc.Fields.Update()

        # 如果启用 eqarr 格式，先保存到临时文件
        if use_eqarr:
            template_path = Path(doc_path).parent / "t1.docx"
            if not template_path.exists():
                print(f"\n警告：未找到模板文件 {template_path}")
                print("将使用标准格式保存")
                use_eqarr = False
            else:
                # 保存到临时文件
                temp_dir = tempfile.mkdtemp()
                temp_file = os.path.join(temp_dir, "temp_numbered.docx")
                doc.SaveAs2(temp_file)
                doc.Close()
                
                print(f"\n正在进行 eqArr 格式转换...")
                if postprocess_eqarr_format(temp_file, str(template_path), str(output_path), simple_mode):
                    print(f"eqArr 格式转换完成：{output_path}")
                    # 清理临时文件
                    try:
                        os.remove(temp_file)
                        os.rmdir(temp_dir)
                    except:
                        pass
                    return 0
                else:
                    print("eqArr 转换失败，使用标准格式")
                    # 回退到标准格式
                    doc = word.Documents.Open(temp_file)
                    try:
                        os.remove(temp_file)
                        os.rmdir(temp_dir)
                    except:
                        pass

        # 标准格式保存
        if Path(output_path).exists():
            Path(output_path).unlink()

        doc.SaveAs2(str(output_path))
        print(f'已保存到：{output_path}')

        # 验证书签
        print(f'\n验证书签:')
        for bm in doc.Bookmarks:
            if bm.Name.startswith('eq'):
                print(f'  - {bm.Name}')

        doc.Close()
        return 0

    except ValueError as e:
        print(f'\n❌ 错误：{e}')
        return 1
    except Exception as e:
        print(f'\n❌ 未知错误：{e}')
        import traceback
        traceback.print_exc()
        return 1
    finally:
        try:
            if doc:
                doc.Close(SaveChanges=False)
        except:
            pass

        try:
            if word:
                word.Quit()
        except:
            pass


def main():
    parser = argparse.ArgumentParser(
        description='DFN - DocxFormulaNumbering | Word 文档公式自动编号工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
示例:
  # 默认模式：使用章节号 - 序号格式 (1-1), (1-2), (2-1)...
  python dfn.py input.docx output.docx

  # 简单模式：使用连续编号格式 (1), (2), (3)...
  python dfn.py input.docx output.docx --simple

  # 使用 t1 格式（eqArr 环境，需要 t1.docx 模板）
  python dfn.py input.docx output.docx --t1

  # 覆盖原文件
  python dfn.py document.docx

格式说明:
  默认/t2 格式：<m:t>E=mc²#(1)</m:t> - 编号为普通文本
  t1 格式：<m:eqArr><m:t>E=mc²#</m:t><m:d>1</m:d></m:eqArr> - 编号在 delimiter 元素中

  t1 格式需要 t1.docx 作为模板文件

前提条件:
  默认模式下，文档需要配置多级列表：
  1. 章节标题使用"标题 1"样式
  2. 通过"开始"→"段落"→"多级列表"应用编号
  3. 标题样式与多级列表级别关联
        '''
    )

    parser.add_argument('input', help='输入 docx 文件路径')
    parser.add_argument('output', nargs='?', default=None, help='输出 docx 文件路径 (可选，默认覆盖原文件)')
    parser.add_argument('--simple', action='store_true', help='使用简单连续编号模式 (不含章节号)')
    parser.add_argument('--t1', action='store_true', help='使用 t1 格式（eqArr 环境，需要 t1.docx 模板）')

    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve() if args.output else None

    if not input_path.exists():
        print(f'错误：文件不存在 - {input_path}')
        return 1

    return add_formula_numbering(
        str(input_path),
        str(output_path) if output_path else None,
        simple_mode=args.simple,
        use_eqarr=args.t1
    )


if __name__ == '__main__':
    import sys
    sys.exit(main())
