# -*- coding: utf-8 -*-
import zipfile
import os
import xml.etree.ElementTree as ET

template_path = r"C:\Users\14362\Desktop\工作\AI agent\农夫模板\农夫山泉2025年10月月报_20251119.pptx"

print("=" * 60)
print("PPT模板分析报告")
print("=" * 60)
print(f"\n文件路径: {template_path}")
print(f"文件存在: {os.path.exists(template_path)}")

if not os.path.exists(template_path):
    print("错误: 文件不存在!")
    exit(1)

file_size = os.path.getsize(template_path)
print(f"文件大小: {file_size / 1024:.2f} KB")

with zipfile.ZipFile(template_path, 'r') as z:
    file_list = z.namelist()
    print(f"\n文件总数: {len(file_list)}")
    
    extensions = {}
    for f in file_list:
        ext = os.path.splitext(f)[1] or 'no-ext'
        extensions[ext] = extensions.get(ext, 0) + 1
    
    print("\n文件类型统计:")
    for ext, count in sorted(extensions.items()):
        print(f"  {ext}: {count}")
    
    slides = [f for f in file_list if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
    print(f"\n幻灯片数量: {len(slides)}")
    
    if slides:
        print("\n" + "=" * 60)
        print("幻灯片结构分析 (前3页)")
        print("=" * 60)
        
        for i, slide_path in enumerate(slides[:3]):
            print(f"\n--- Slide {i+1}: {slide_path} ---")
            try:
                with z.open(slide_path) as slide_file:
                    content = slide_file.read().decode('utf-8')
                    root = ET.fromstring(content)
                    
                    ns = {
                        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
                        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'
                    }
                    
                    shapes = root.findall('.//p:sp', ns)
                    pics = root.findall('.//p:pic', ns)
                    texts = root.findall('.//a:t', ns)
                    
                    print(f"  形状: {len(shapes)}, 图片: {len(pics)}, 文本: {len(texts)}")
                    
                    if texts:
                        print("  文本示例:")
                        for j, text in enumerate(texts[:8]):
                            if text.text and text.text.strip():
                                t = text.text.strip()[:60]
                                print(f"    {j+1}. {t}")
                                
            except Exception as e:
                print(f"  Error: {e}")
    
    layouts = [f for f in file_list if f.startswith('ppt/slideLayouts/') and f.endswith('.xml')]
    masters = [f for f in file_list if f.startswith('ppt/slideMasters/') and f.endswith('.xml')]
    print(f"\n布局数量: {len(layouts)}, 母版数量: {len(masters)}")

print("\n" + "=" * 60)
print("Done!")
print("=" * 60)