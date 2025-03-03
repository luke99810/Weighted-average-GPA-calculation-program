import pandas as pd
import os

# 设定GPA计算规则
def convert_grade(grade):
    if pd.isna(grade):
        return None
    if isinstance(grade, str):
        grade = grade.strip()
        if grade == '优秀':
            return 4.0
        elif grade == '良好':
            return 3.3
        elif grade == '中等':
            return 2.3
        elif grade == '及格':
            return 1.3
        elif grade == '不及格':
            return 0.0
        else:
            try:
                return convert_numeric_grade(float(grade))
            except:
                return None
    elif isinstance(grade, (int, float)):
        return convert_numeric_grade(float(grade))
    else:
        return None


def convert_numeric_grade(score):
    if score >= 90:
        return 4.0
    elif score >= 87:
        return 3.7
    elif score >= 84:
        return 3.3
    elif score >= 80:
        return 3.0
    elif score >= 77:
        return 2.7
    elif score >= 74:
        return 2.3
    elif score >= 70:
        return 2.0
    elif score >= 67:
        return 1.7
    elif score >= 64:
        return 1.3
    elif score >= 60:
        return 1.0
    else:
        return 0.0


# 读取Excel文件
file_path = r"C:\Users\你的原始成绩文件.xlsx"
df = pd.read_excel(file_path, sheet_name='sheet1', header=0)

# 提取课程信息（从第4列开始到倒数第二列，根据实际数据修改）
course_columns = df.columns[3:-1]

# 解析学分（列名格式为课程号-课程名称-学分）
credits = []
for col in course_columns:
    parts = col.split('-')
    if len(parts) >= 3:
        credit_str = parts[-1]
        try:
            credit = float(credit_str)
            credits.append(credit)
        except:
            credits.append(0.0)
    else:
        credits.append(0.0)

# 计算每个学生的平均学分绩点
results = []
for index, row in df.iterrows():
    total_points = 0.0
    total_credits = 0.0
    for i, col in enumerate(course_columns):
        credit = credits[i]
        grade = row[col]
        grade_point = convert_grade(grade)
        if grade_point is not None and credit > 0:
            total_points += grade_point * credit
            total_credits += credit
    gpa = total_points / total_credits if total_credits != 0 else 0.0
    results.append({
        '姓名': row['姓名'],
        '班级': row['班级'],
        '计算平均学分绩点': round(gpa, 4)
    })

# 转换为DataFrame输出
result_df = pd.DataFrame(results)
print("计算结果：")
print(result_df.to_string(index=False))

# 自动保存到指定桌面路径
default_save_path = r"C:\Users\加权平均成绩表.xlsx"

try:
    # 自动创建目录（如果不存在）
    os.makedirs(os.path.dirname(default_save_path), exist_ok=True)
    result_df.to_excel(default_save_path, index=False)
    print(f"\n✅ 结果已自动保存到：{os.path.abspath(default_save_path)}")
except Exception as e:
    print(f"\n❌ 自动保存失败：{str(e)}")

# 可选：用户自定义保存路径
custom_path = input("\n需要保存到其他位置吗？直接回车跳过，或输入完整路径：").strip()

if custom_path:
    try:
        # 处理扩展名和路径
        if not custom_path.lower().endswith('.xlsx'):
            custom_path += '.xlsx'

        # 创建目录（如果不存在）
        os.makedirs(os.path.dirname(custom_path), exist_ok=True)

        # 保存文件
        result_df.to_excel(custom_path, index=False)
        print(f"✅ 结果已保存到：{os.path.abspath(custom_path)}")
    except Exception as e:
        print(f"❌ 保存到自定义路径失败：{str(e)}")