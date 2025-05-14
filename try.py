from werkzeug.utils import secure_filename
print(secure_filename('Смета.xlsx'))
print(secure_filename('Абв.txt'))
print(secure_filename('Test File.zip'))