import os
from paddleocr import PaddleOCR

def ocr_image_folder():
    # 初始化 PaddleOCR
    folder_path='RedsealRomove/remove_res'
    ocr = PaddleOCR(use_gpu=False,
                    lang="ch",
                    enable_mkldnn=True,
                    det_model_dir="whl/det/ch/ch_PP-OCRv4_det_infer/",
                    cls_model_dir="whl/cls/ch_ppocr-mobile_v2.0_cls_infer/",
                    rec_model_dir="whl/rec/ch/ch_PP-OCRv4_rec_infer/"
                    )

    # 创建用于存放txt的文件夹
    txt_folder_path = os.path.join(os.path.dirname(folder_path), 'txt')
    os.makedirs(txt_folder_path, exist_ok=True)

    # 遍历文件夹中的所有图片文件
    for file_name in sorted(os.listdir(folder_path), key=lambda x: int(os.path.splitext(x)[0])):
        if file_name.endswith(('.jpg', '.jpeg', '.png')):  # 只处理图片文件
            file_path = os.path.join(folder_path, file_name)

            # 使用 PaddleOCR 进行图像识别
            result = ocr.ocr(file_path, cls=True)
            #print(result)
            # 将识别结果写入txt文件
            txt_file_path = os.path.join(txt_folder_path, os.path.splitext(file_name)[0] + '.txt')
            with open(txt_file_path, 'w', encoding='utf-8') as f:
                for line in result:
                    line_text = ' '.join([str(word_info[1][0]) for word_info in line])

                    #print(line_text)
                    f.write(line_text + '\n')

            print(f'完成图片 {file_name} 的识别，生成了 {txt_file_path}')


# 调用函数进行图片识别和生成txt文件
def run_3():
    ocr_image_folder()
