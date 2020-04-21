import os
import re
import time

src_file = r'F:\123.txt'
dst_file = r'F:\555.txt'

src_dir = r'F:\中国通史'
dst_dir = r'F:\甄嬛传'

def modify_file():
    try:
        os.rename(src_file, dst_file)
    except Exception as e:
        print(e)
        print('rename file fail\r\n')
    else:
        print('rename file success\r\n')

def modify_dir():
    try:
        os.rename(src_dir, dst_dir)
    except Exception as e:
        print(e)
        print('rename directory fail\r\n')
    else:
        print('rename directory success\r\n')

def modify_files():
    for each_file in os.listdir(src_dir):
        result_num = re.search(r'EP[0][1-9]\.', each_file)
        if result_num:
            print(result_num.group())
            result_file_format = re.findall(r'\.\w+', each_file)
            file_num = 'EP0' + result_num.group()[2] + result_num.group()[3]
            new_filename = '[中国通史].General.History.of.China.' + file_num + '.2013.HDTV.720p.x264.AC3' + result_file_format[-1]
            try:
                print(each_file)
                print(new_filename)
                old_file = src_dir + '\\' + each_file
                new_file = src_dir + '\\' + new_filename
                # # print(old_file)
                # # print(new_file)
                os.rename(old_file, new_file)
            except Exception as e:
                print(e)
                print('rename file fail\r\n')
                time.sleep(10)
            else:
                print('rename file success\r\n')


if __name__ == "__main__":
    # modify_file()
    # modify_dir()
    modify_files()
