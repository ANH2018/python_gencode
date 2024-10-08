import pandas as pd

# Đọc tệp Excel
file_path = 'A.xlsx'
df = pd.read_excel(file_path)

# Hàm để làm sạch và chuyển đổi giá trị
def process_values(values):
    # Giữ nguyên vị trí đầu tiên, các vị trí còn lại nhân 1000
    return [values[0]] + [value * 1000 for value in values[1:]]

# Hàm in ra mã theo yêu cầu và ghi vào file
def generate_code(values, file):
    # Chuyển đổi giá trị theo quy tắc
    transformed_values = process_values(values)
    leng = len(transformed_values)
    
    # Ghi ra mã theo định dạng yêu cầu vào file
    file.write(f"// {transformed_values} leng {leng}\n")
    file.write(f"if ( effect == {transformed_values[0]})\n")
    file.write("{\n")
    
    # Ghi TURNON_LED và TURNOFF_LED xen kẽ với delay()
    for i in range(1, leng, 2):  # Lặp qua từng cặp giá trị
        file.write(" TURNON_LED();\n")
        file.write(f" delay({transformed_values[i]});\n")  # Ghi delay tại vị trí chẵn
        if i + 1 < leng:  # Kiểm tra nếu có vị trí tiếp theo
            file.write(" TURNOFF_LED();\n")
            file.write(f" delay({transformed_values[i + 1]});\n")  # Ghi delay tại vị trí lẻ
        else:
            file.write(" TURNOFF_LED();\n")  # Ghi TURNOFF_LED cuối cùng nếu không có giá trị lẻ tiếp theo

    file.write("}\n\n")  # Đóng dấu ngoặc nhọn và xuống dòng

# Mở tệp để ghi dữ liệu
with open('code.txt', 'w') as file:
    # Duyệt qua từng dòng trong tệp Excel và tạo mã
    for index, row in df.iterrows():
        # Làm sạch dữ liệu: loại bỏ các giá trị NaN và giá trị bằng 0
        cleaned_values = [value for value in row if pd.notna(value) and value != 0]
        if cleaned_values:
            generate_code(cleaned_values, file)

print("Các lệnh đã được ghi vào file code.txt.")
