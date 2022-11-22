import os
import winshell
import sys
import win32com.client

SOURCE = "G:/work/' BoOks/Книга об о. Димитрии/ФОТО"
DESTINATION_DISK = 'C:'
DESTINATION_PATH = '/temp/'
DEST = os.path.join(DESTINATION_DISK, DESTINATION_PATH)
destination = os.path.join(DESTINATION_DISK, DESTINATION_PATH)


def file_copy_func(source, destination):
    if os.path.exists(destination):
        return
    if not os.path.exists(os.path.split(destination)[0]):
        os.makedirs(os.path.split(destination)[0])
    command = 'copy ' + '"' + source + '"' + ' ' + '"' + os.path.normpath(destination) + '"'
    print(command)
    os.system(command)





def file_copy(source, destination):
    for item in os.listdir(source):
        # print(item)
        # print(os.path.isfile(os.path.join(source, item)))
        if os.path.isfile(os.path.join(source, item)):
            # print(os.path.splitext(os.path.join(source, item)))
            if os.path.splitext(os.path.join(source, item))[1] != '.lnk':
                file_copy_func(os.path.join(source, item), os.path.join(destination, item))
                # print(os.path.join(source, item),' -- ', os.path.join(destination, item))  # TODO заменить на копирование файла
                continue
            shortcut = winshell.shortcut(os.path.join(source, item))
            # print(os.path.abspath(item))
            # print(shortcut.path)
            if os.path.isfile(shortcut.path):
                file_copy_func(shortcut.path, os.path.join(destination, os.path.split(shortcut.path)[1]))
                # print(shortcut.path, ' -- ', os.path.join(destination, os.path.split(shortcut.path)[1]))  # TODO заменить на копирование файла
                continue
            if os.path.isdir(shortcut.path):
                if destination == DEST:
                    file_copy(shortcut.path, os.path.join(destination, os.path.split(shortcut.path)[1]))
                    continue
                file_copy(shortcut.path, os.path.join(destination))
                continue
        if os.path.isdir(os.path.join(source, item)):
            if destination == DEST:
                file_copy(os.path.join(source, item), os.path.join(destination, item))
            else:
                file_copy(os.path.join(source, item), os.path.join(destination))


def main():
    file_copy(SOURCE, destination)


if __name__ == '__main__':
    main()