# OperateExcel
    从指定Excel列后的某一行开始复制到另一个Excel指定列后的行

# Version
    VersioVersion：1.0.2

# next VersionT arget

    1，目前从复制到目的存在问题，创建目标表无法正常copy格式，导致超行
    2，优化复制过程，使之可以创建新文件复制。
    3，批量文件夹内文档复制到目标文件夹，并更改目标文件名。


# Last commit note

NoteBook

Version：1.0.2

date：23/10/14 22:59

des：修复了可能会出现Zip bomb detected!的bug，将第3行复制到第7行的逻辑改变

    ZipSecureFile.setMinInflateRatio(-1.0d);

BUG：目前使用setRowCellData方法将第三行复制到第七行有问题。

nextVersionTarget：

    1，目前从复制到目的存在问题，创建目标表无法正常copy格式，导致超行
    2，优化复制过程，使之可以创建新文件复制。
    3，批量文件夹内文档复制到目标文件夹，并更改目标文件名。
