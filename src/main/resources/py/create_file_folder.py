import os

# 创建file文件夹
current_dir = os.path.dirname(os.path.abspath(__file__))
file_dir = os.path.join(current_dir, 'file')

if not os.path.exists(file_dir):
    os.makedirs(file_dir)
    print("'file'文件夹已创建")
else:
    print("'file'文件夹已存在") 