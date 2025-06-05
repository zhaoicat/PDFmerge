#!/usr/bin/env python3
"""
PDF交替合并工具
按页面生成多个PDF文件：
- 结果1.pdf: A的第1页 + B的第1页 + C的第1页
- 结果2.pdf: A的第2页 + B的第2页 + C的第2页
- 以此类推...
"""

import PyPDF2
import os
import sys
import time
import re
from pathlib import Path
from tqdm import tqdm

# 设置Poppler路径
def setup_poppler_path():
    """自动设置Poppler路径"""
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        current_dir = Path(sys._MEIPASS)
    else:
        # 如果是开发环境
        current_dir = Path(__file__).parent
    
    poppler_paths = [
        current_dir / "poppler" / "poppler-24.08.0" / "Library" / "bin",
        current_dir / "poppler" / "bin",
        Path("C:/poppler/bin"),
        Path("C:/Program Files/poppler/bin"),
    ]
    
    for poppler_path in poppler_paths:
        if poppler_path.exists():
            poppler_str = str(poppler_path.absolute())
            current_path = os.environ.get('PATH', '')
            if poppler_str not in current_path:
                os.environ['PATH'] = poppler_str + os.pathsep + current_path
                print(f"🔧 已设置Poppler路径: {poppler_str}")
            return True
    
    print("⚠️ 未找到Poppler，OCR功能可能不可用")
    return False

# 在导入OCR库之前设置Poppler路径
setup_poppler_path()

# OCR相关导入
try:
    import pytesseract
    from pdf2image import convert_from_path
    
    # 设置Tesseract路径和语言包路径
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe
        current_dir = Path(sys._MEIPASS)
        tesseract_exe = current_dir / "tesseract" / "tesseract.exe"
        tessdata_dir = current_dir / "tesseract" / "tessdata"
    else:
        # 如果是开发环境
        current_dir = Path(__file__).parent
        tesseract_exe = None
        tessdata_dir = current_dir / "tesseract" / "tessdata"
    
    # 设置TESSDATA_PREFIX环境变量
    if tessdata_dir.exists():
        os.environ['TESSDATA_PREFIX'] = str(tessdata_dir)
        print(f"🔧 设置TESSDATA_PREFIX: {tessdata_dir}")
    
    tesseract_paths = []
    if tesseract_exe and tesseract_exe.exists():
        tesseract_paths.append(str(tesseract_exe))
    tesseract_paths.extend([
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Tesseract-OCR\tesseract.exe",
        "tesseract"  # 如果已在PATH中
    ])
    
    tesseract_found = False
    for path in tesseract_paths:
        try:
            pytesseract.pytesseract.tesseract_cmd = path
            # 测试是否可以运行
            pytesseract.get_tesseract_version()
            tesseract_found = True
            print(f"✅ 找到Tesseract: {path}")
            break
        except:
            continue
    
    if not tesseract_found:
        raise ImportError("Tesseract not found")
    
    OCR_AVAILABLE = True
    print("✅ OCR功能已启用")
except ImportError:
    OCR_AVAILABLE = False
    msg = "❌ OCR功能不可用，需要安装: pip install pytesseract pdf2image Pillow"
    print(msg)
    print("❌ 或者需要安装Tesseract OCR引擎")
    print("❌ 程序要求必须有OCR功能，现在退出...")
    exit(1)


def extract_text_with_ocr(pdf_path, page_num=0):
    """
    使用OCR从PDF页面提取文字
    
    Args:
        pdf_path: PDF文件路径
        page_num: 页面索引（从0开始）
    
    Returns:
        str: 提取的文字，失败返回None
    """
    if not OCR_AVAILABLE:
        print("❌ OCR功能不可用，程序退出")
        exit(1)
    
    try:
        print(f"  🔍 使用OCR提取第{page_num + 1}页文字...")
        
        # 将PDF页面转换为图像
        images = convert_from_path(pdf_path, first_page=page_num + 1,
                                   last_page=page_num + 1, dpi=300)
        
        if images:
            # 使用OCR提取文字
            text = pytesseract.image_to_string(images[0], lang='chi_sim+eng')
            print(f"  📝 OCR提取文字长度: {len(text)} 字符")
            # 显示提取的文字内容（用于调试）
            if text.strip():
                print(f"  📄 OCR提取的文字内容:")
                print(f"  {repr(text[:500])}")  # 显示前500个字符
            return text
        else:
            print("  ❌ 无法转换PDF页面为图像")
            return None
            
    except Exception as e:
        error_msg = str(e)
        print(f"  ❌ OCR提取失败: {error_msg}")
        # 如果是tesseract相关错误，直接退出程序
        if "tesseract" in error_msg.lower():
            print("❌ Tesseract OCR引擎未安装或未配置，程序退出")
            print("❌ 请安装Tesseract OCR: https://github.com/tesseract-ocr/tesseract")
            exit(1)
        return None


def extract_applicant_name(pdf_reader, pdf_path):
    """
    从PDF中提取申请人姓名（"申请人："和"，"之间的文字）
    先尝试直接文本提取，如果失败则使用OCR
    
    Args:
        pdf_reader: PDF阅读器对象
        pdf_path: PDF文件路径（用于日志）
    
    Returns:
        str: 申请人姓名，如果未找到则返回None
    """
    try:
        # 方法1：直接文本提取
        print("  📖 尝试直接文本提取...")
        for page_num, page in enumerate(pdf_reader.pages):
            text = page.extract_text()
            if text and text.strip():
                # 使用正则表达式查找"申请人："和"，"之间的文字
                patterns = [
                    r'申请人：([^，]+)，',
                    r'申请人:([^，]+)，',
                    r'申请人：([^,]+),',
                    r'申请人:([^,]+),',
                    r'申\s*请\s*人\s*[:：]\s*([^，,\s]+)',  # 处理空格和变体
                    r'申\s*请\s*人\s*[:：]\s*([^，,]+?)\s*[，,]',  # 更宽松的匹配
                ]
                for pattern in patterns:
                    match = re.search(pattern, text)
                    if match:
                        applicant_name = match.group(1).strip()
                        # 去掉姓名中的空格
                        applicant_name = re.sub(r'\s+', '', applicant_name)
                        print(f"  📝 在第{page_num + 1}页找到申请人: {applicant_name}")
                        return applicant_name
        
        print("  ⚠️  直接文本提取未找到申请人信息")
        
        # 方法2：OCR提取
        if OCR_AVAILABLE:
            print("  🔍 尝试OCR文本提取...")
            for page_num in range(len(pdf_reader.pages)):
                ocr_text = extract_text_with_ocr(pdf_path, page_num)
                if ocr_text:
                    # 使用正则表达式查找"申请人："和"，"之间的文字
                    patterns = [
                        r'申请人：([^，]+)，',
                        r'申请人:([^，]+)，',
                        r'申请人：([^,]+),',
                        r'申请人:([^,]+),',
                        r'申\s*请\s*人\s*[:：]\s*([^，,男女]+?)\s*[，,男女]',  # 匹配到逗号或性别字符
                        r'申\s*请\s*人\s*[:：]\s*([^\s，,]+(?:\s+[^\s，,]+)*?)\s*[，,男女]',  # 处理姓名中的空格
                        r'申\s*请\s*人\s*[:：]\s*([^，,\n]+?)\s*[，,]',  # 更宽松的匹配
                    ]
                    
                    for pattern in patterns:
                        match = re.search(pattern, ocr_text)
                        if match:
                            applicant_name = match.group(1).strip()
                            # 去掉姓名中的空格
                            applicant_name = re.sub(r'\s+', '', applicant_name)
                            print(f"  📝 OCR在第{page_num + 1}页找到申请人: {applicant_name}")
                            return applicant_name
            
            print("  ⚠️  OCR也未找到申请人信息")
        else:
            print("❌ OCR功能不可用，程序退出")
            exit(1)
        
        print(f"  ❌ 在 {os.path.basename(pdf_path)} 中未找到申请人信息")
        return None
        
    except Exception as e:
        print(f"  ❌ 提取申请人信息时出错: {e}")
        return None


def merge_pdfs_by_page_groups(pdf_paths, output_dir):
    """
    按页面组合并多个PDF文件，生成多个结果文件
    
    Args:
        pdf_paths (list): PDF文件路径列表
        output_dir (str): 输出目录路径
    """
    start_time = time.time()
    start_time_str = time.strftime('%Y-%m-%d %H:%M:%S', 
                                   time.localtime(start_time))
    print(f"⏰ 开始处理时间: {start_time_str}")
    
    # 验证文件存在性
    print("🔍 验证文件存在性...")
    for pdf_path in pdf_paths:
        if not os.path.exists(pdf_path):
            print(f"❌ 文件不存在: {pdf_path}")
            return False
    print("✅ 所有文件验证通过")
    
    # 打开所有PDF文件并保持文件句柄
    pdf_files = []
    pdf_readers = []
    
    try:
        # 打开所有文件
        print("📖 正在读取PDF文件...")
        for pdf_path in tqdm(pdf_paths, desc="读取文件", unit="文件"):
            file_handle = open(pdf_path, 'rb')
            pdf_files.append(file_handle)
            reader = PyPDF2.PdfReader(file_handle)
            pdf_readers.append((reader, pdf_path))
        
        # 从第3个PDF的每一页分别提取申请人姓名
        page_applicant_names = []
        if len(pdf_readers) >= 3:
            print("📝 正在从第3个PDF的每一页提取申请人信息...")
            third_reader, third_path = pdf_readers[2]  # 第三个PDF（索引2）
            pdf_name = os.path.basename(third_path)
            print(f"🔍 处理文件: {pdf_name}")
            
            # 为每一页分别提取申请人姓名
            for page_num in range(len(third_reader.pages)):
                print(f"\n📄 处理第{page_num + 1}页...")
                
                # 方法1：直接文本提取
                page = third_reader.pages[page_num]
                text = page.extract_text()
                applicant_name = None
                
                if text and text.strip():
                    print("  📖 尝试直接文本提取...")
                    patterns = [
                        r'申请人：([^，]+)，',
                        r'申请人:([^，]+)，',
                        r'申请人：([^,]+),',
                        r'申请人:([^,]+),',
                        r'申\s*请\s*人\s*[:：]\s*([^，,男女]+?)\s*[，,男女]',  # 匹配到逗号或性别字符
                        r'申\s*请\s*人\s*[:：]\s*([^\s，,]+(?:\s+[^\s，,]+)*?)\s*[，,男女]',  # 处理姓名中的空格
                        r'申\s*请\s*人\s*[:：]\s*([^，,\n]+?)\s*[，,]',  # 更宽松的匹配
                    ]
                    applicant_name = None
                    for pattern in patterns:
                        match = re.search(pattern, text)
                        if match:
                            applicant_name = match.group(1).strip()
                            # 去掉姓名中的空格
                            applicant_name = re.sub(r'\s+', '', applicant_name)
                            print(f"  📝 找到申请人: {applicant_name}")
                            break
                
                # 方法2：OCR提取（如果直接提取失败）
                if not applicant_name:
                    if OCR_AVAILABLE:
                        print("  🔍 尝试OCR文本提取...")
                        ocr_text = extract_text_with_ocr(third_path, page_num)
                        if ocr_text:
                            patterns = [
                                r'申请人：([^，]+)，',
                                r'申请人:([^，]+)，',
                                r'申请人：([^,]+),',
                                r'申请人:([^,]+),',
                                r'申\s*请\s*人\s*[:：]\s*([^，,男女]+?)\s*[，,男女]',  # 匹配到逗号或性别字符
                                r'申\s*请\s*人\s*[:：]\s*([^\s，,]+(?:\s+[^\s，,]+)*?)\s*[，,男女]',  # 处理姓名中的空格
                                r'申\s*请\s*人\s*[:：]\s*([^，,\n]+?)\s*[，,]',  # 更宽松的匹配
                            ]
                            
                            for pattern in patterns:
                                match = re.search(pattern, ocr_text)
                                if match:
                                    applicant_name = match.group(1).strip()
                                    # 去掉姓名中的空格
                                    applicant_name = re.sub(r'\s+', '', applicant_name)
                                    print(f"  📝 OCR找到申请人: {applicant_name}")
                                    break
                    else:
                        print("❌ OCR功能不可用，程序退出")
                        exit(1)
                
                page_applicant_names.append(applicant_name)
                
                if applicant_name:
                    print(f"  ✅ 第{page_num + 1}页申请人: {applicant_name}")
                else:
                    print(f"  ❌ 第{page_num + 1}页未找到申请人")
        else:
            print("⚠️  PDF文件少于3个，无法提取申请人信息")
        
        print("\n📊 每页申请人姓名提取结果:")
        for i, name in enumerate(page_applicant_names):
            status = "✅ {}".format(name) if name else "❌ 未提取"
            print("  第{}页: {}".format(i+1, status))
        
        # 获取最大页数
        max_pages = max(len(reader.pages) for reader, _ in pdf_readers)
        print(f"\n📄 最大页数: {max_pages}")
        print(f"🎯 将生成 {max_pages} 个结果PDF文件")
        
        generated_files = []
        
        # 为每一页生成一个独立的PDF文件
        print("\n🔄 开始生成分页合并文件...")
        progress_bar = tqdm(range(max_pages), desc="生成PDF", unit="页")
        
        for page_num in progress_bar:
            page_start_time = time.time()
            pdf_writer = PyPDF2.PdfWriter()
            current_page_files = []
            
            # 更新进度条描述
            progress_bar.set_description(f"处理第{page_num + 1}页")
            
            # 从每个PDF中取出对应页面
            for i, (reader, pdf_path) in enumerate(pdf_readers):
                if page_num < len(reader.pages):
                    page = reader.pages[page_num]
                    pdf_writer.add_page(page)
                    pdf_name = os.path.basename(pdf_path)
                    current_page_files.append(pdf_name)
                else:
                    pdf_name = os.path.basename(pdf_path)
                    tqdm.write(f"  ⚠️  PDF {pdf_name} 没有第{page_num + 1}页，跳过")
            
            # 生成输出文件名
            if page_num < len(page_applicant_names) and page_applicant_names[page_num]:
                applicant_name = page_applicant_names[page_num]
                output_filename = f"{applicant_name}.pdf"
            else:
                output_filename = f"第{page_num + 1}页合并结果.pdf"
            
            output_path = os.path.join(output_dir, output_filename)
            
            # 写入文件
            with open(output_path, 'wb') as output_file:
                pdf_writer.write(output_file)
            
            generated_files.append(output_path)
            
            # 计算单页处理时间
            page_time = time.time() - page_start_time
            files_str = ', '.join(current_page_files)
            time_info = f"耗时: {page_time:.2f}秒"
            msg = f"  ✅ 生成: {output_filename} (包含: {files_str}) - {time_info}"
            tqdm.write(msg)
        
        # 计算总处理时间
        total_time = time.time() - start_time
        end_time = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())
        
        print(f"\n🎉 所有文件生成完成! 共生成 {len(generated_files)} 个PDF文件")
        print(f"⏰ 完成时间: {end_time}")
        print(f"⏱️  总耗时: {total_time:.2f} 秒")
        print(f"📊 平均每页处理时间: {total_time/max_pages:.2f} 秒")
        return True
        
    except Exception as e:
        error_time = time.time() - start_time
        print(f"\n❌ 处理过程中出错: {e}")
        print(f"⏱️  错误前耗时: {error_time:.2f} 秒")
        return False
    
    finally:
        # 关闭所有文件句柄
        print("🔒 正在关闭文件...")
        for file_handle in pdf_files:
            try:
                file_handle.close()
            except Exception:
                pass


def main():
    """主函数"""
    print("🔥 PDF分页合并工具 (需要OCR)")
    print("=" * 50)
    print("📝 功能：将多个PDF按页面生成独立的合并文件")
    print("💡 文件命名：使用第三个PDF中申请人姓名命名")
    if OCR_AVAILABLE:
        print("🔍 OCR功能：已启用，支持扫描版PDF文字提取")
    else:
        print("❌ OCR功能：未启用，程序无法运行")
        print("❌ 请安装OCR依赖: pip install pytesseract pdf2image Pillow")
        print("❌ 程序退出")
        exit(1)
    print("=" * 50)
    
    # 记录程序开始时间
    program_start_time = time.time()
    
    # 获取正确的工作目录
    if getattr(sys, 'frozen', False):
        # 如果是打包后的exe，使用exe文件所在目录的上级目录（项目根目录）
        exe_dir = Path(sys.executable).parent
        if exe_dir.name == 'dist':
            # 如果exe在dist目录中，使用dist的上级目录
            current_dir = exe_dir.parent
        else:
            # 如果exe直接在项目根目录，使用exe所在目录
            current_dir = exe_dir
    else:
        # 如果是开发环境，使用脚本文件所在目录
        current_dir = Path(__file__).parent
    
    # 设置正确的路径：原始文件在子目录PDF插入下
    pdf_dir = current_dir / "PDF插入" / "原始文件"
    output_dir = current_dir / "PDF插入" / "最终文件"
    
    # 查找PDF文件
    print("🔍 扫描PDF文件...")
    pdf_files = list(pdf_dir.glob("*.pdf"))
    pdf_files.sort()  # 按名称排序
    
    if not pdf_files:
        print("❌ 在原始文件目录中没有找到PDF文件")
        print(f"查找目录: {pdf_dir}")
        return
    
    print("找到{}个PDF文件:".format(len(pdf_files)))
    for i, pdf_file in enumerate(pdf_files, 1):
        # 获取文件大小
        file_size = pdf_file.stat().st_size / 1024 / 1024  # MB
        marker = " 📝(提取申请人)" if i == 3 else ""
        print(f"  {i}. {pdf_file.name} ({file_size:.1f} MB){marker}")
    
    # 设置输出路径
    output_dir.mkdir(exist_ok=True)
    print(f"📁 输出目录: {output_dir}")
    
    # 执行合并
    print(f"\n🚀 开始处理 {len(pdf_files)} 个PDF文件...")
    pdf_paths = [str(pdf_file) for pdf_file in pdf_files]
    
    success = merge_pdfs_by_page_groups(pdf_paths, str(output_dir))
    
    # 计算程序总运行时间
    program_total_time = time.time() - program_start_time
    
    if success:
        print("\n🎊 处理完成!")
        print("输出目录: {}".format(output_dir))
        print(f"🕒 程序总运行时间: {program_total_time:.2f} 秒")
    else:
        print(f"\n💥 处理失败! 总耗时: {program_total_time:.2f} 秒")


if __name__ == "__main__":
    main() 