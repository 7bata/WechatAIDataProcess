import os
import glob


def rename_files_in_folder(folder_path):
    # Get all files in the target folder
    files = glob.glob(os.path.join(folder_path, '*'))

    for index, file_path in enumerate(files):
        # Get the file extension
        ext = os.path.splitext(file_path)[1]

        # new file name
        new_name = f"file_{index + 1}{ext}"
        new_file_path = os.path.join(folder_path, new_name)

        # rename
        os.rename(file_path, new_file_path)
        print(f"Finished: {file_path} -> {new_file_path}")


# target folder path
folder_path = 'D:\WechatAI\WebContent\PDF'
rename_files_in_folder(folder_path)