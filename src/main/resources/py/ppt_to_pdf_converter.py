import os
from pathlib import Path
import comtypes.client

def convert_pptx_to_pdf():
    # 获取当前目录下的file文件夹路径
    current_dir = os.path.dirname(os.path.abspath(__file__))
    file_dir = os.path.join(current_dir, 'file')
    
    if not os.path.exists(file_dir):
        print("Error: 'file'文件夹不存在！")
        return

    # PowerPoint常量
    POWERPOINT_FORMATTYPE_PDF = 32

    try:
        # 创建PowerPoint应用程序实例
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  # 可以设置为0来隐藏界面

        # 遍历所有的.pptx文件
        for pptx_file in Path(file_dir).glob('*.pptx'):
            try:
                pptx_path = str(pptx_file.absolute())
                pdf_path = str(pptx_file.with_suffix('.pdf').absolute())
                
                print(f"正在处理: {pptx_file.name}")
                
                # 打开PPT文件
                presentation = powerpoint.Presentations.Open(pptx_path)
                
                # 转换为PDF
                presentation.SaveAs(pdf_path, POWERPOINT_FORMATTYPE_PDF)
                
                # 关闭演示文稿
                presentation.Close()
                
                print(f"转换完成: {pptx_file.name} -> {Path(pdf_path).name}")
                
            except Exception as e:
                print(f"处理文件 {pptx_file.name} 时发生错误: {str(e)}")

    except Exception as e:
        print(f"PowerPoint 应用程序错误: {str(e)}")
        
    finally:
        try:
            # 退出PowerPoint应用程序
            powerpoint.Quit()
        except:
            pass

    print("所有文件处理完成！")

if __name__ == "__main__":
    convert_pptx_to_pdf() 