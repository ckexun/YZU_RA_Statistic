import configparser
import pandas as pd

config = configparser.ConfigParser()
config.read('config.ini')
excelfile_with_score = config['ExcelFile']['Score']
excelfile_with_admission = config['ExcelFile']['AdmissionChannels']

# load dataframe
Score_df = pd.read_excel(f'./data/{excelfile_with_score}.xls', header=None)
# Set Score columns name
Score_df.columns = Score_df.iloc[2]
Score_df = Score_df.drop(index=[0,1,2,3]).reset_index(drop=True) # Delete
# Filter 停休、休學...
Score_df = Score_df[["學號", "姓名", "總成績"]]  
Score_df = Score_df.dropna(subset=["總成績"])  # remove NaN
Score_df["總成績"] = Score_df["總成績"].astype(int)

# load dataframe
Admission_df = pd.read_excel(f'./data/{excelfile_with_admission}.xls', header=None)
# Set Score columns name
Admission_df.columns = Admission_df.iloc[7]
Admission_df = Admission_df.drop(index=list(range(8))).reset_index(drop=True) # Delete
Admission_df = Admission_df[["學號", "入學管道"]]

print(Score_df.to_string())
print(Admission_df.to_string())

merged_df = pd.merge(Score_df, Admission_df, on="學號", how="inner")
merged_df.insert(0, "學期", config['ClassInfo']['Semester']) 
merged_df.insert(1, "課號", config['ClassInfo']['ClassID']) 
merged_df.insert(2, "班別", config['ClassInfo']['ClassType']) 
merged_df.insert(3, "課名", config['ClassInfo']['ClassName']) 

# add 3 columns
merged_df["名次"] = merged_df["總成績"].rank(ascending=False, method="min").astype(int)
merged_df["及格"] = merged_df["總成績"].apply(lambda x: "Y" if x >= 60 else "")
merged_df = merged_df.sort_values(by="名次").reset_index(drop=True)
threshold = len(merged_df)*0.3
threshold_int = int(threshold)
threshold = threshold_int if threshold-threshold_int < 0.5 else threshold_int+1
merged_df["前30%"] = ["Y" if i < threshold else "" for i in range(len(merged_df))]
print(merged_df.to_string())


def calculate_admission_statistics(df):
    result = []

    # 遍歷每個入學管道
    for channel, group in df.groupby("入學管道"):
        students_num = len(group)  
        pass_num = group[group["及格"] == "Y"].shape[0]  
        top30_num = group[group["前30%"] == "Y"].shape[0]  
        rank_sum = group["名次"].sum()  

        pass_rate = (pass_num/students_num)*100 
        top30_rate = (top30_num/students_num)*100
        rank_rate = (rank_sum/students_num)/len(df) * 100

        result.append({
            "入學管道": channel,
            "學生人數(不包含停修、休學、退學)": students_num,
            "該科及格人次比例": f"{pass_rate:.2f}%",
            "該科排名前30%人次比例": f"{top30_rate:.2f}%",
            "該科期末成績排名百分比例": f"{rank_rate:.2f}%"
        })

    # 全班
    total_pass_num = df[df["及格"] == "Y"].shape[0]
    total_pass_rate = total_pass_num/len(df)*100
    result.append({
            "入學管道": "全班",
            "學生人數(不包含停修、休學、退學)": len(df),
            "該科及格人次比例": f"{total_pass_rate:.2f}%",
            "該科排名前30%人次比例": "",
            "該科期末成績排名百分比例": ""
        })
    
    result_df = pd.DataFrame(result)

    # Sort by "學生人數(不包含停修、休學、退學)"
    result_df = pd.concat([
        result_df[result_df["入學管道"] != "全班"].sort_values(by="學生人數(不包含停修、休學、退學)", ascending=False),
        result_df[result_df["入學管道"] == "全班"]
    ]).reset_index(drop=True)

    return result_df

sheet1_df = calculate_admission_statistics(merged_df)

print(sheet1_df.to_string())

sheet2_df = merged_df[["學期","課號","班別","課名","入學管道","總成績","名次","及格","前30%"]]
sheet2_df = sheet2_df.rename(columns={"總成績":"分數"})

additional_data = sheet2_df.copy()
additional_data["入學管道"] = "全班"
# merge oringal df and addional df
sheet3_df = pd.concat([sheet2_df, additional_data], ignore_index=True)
sheet3_df = sheet3_df[["學期", "課號", "班別", "課名", "入學管道", "分數"]]
sheet3_df

output_file = f"{config['ClassInfo']['Semester']}-{config['ClassInfo']['ClassID']}{config['ClassInfo']['ClassType']}{config['ClassInfo']['ClassName']}.xlsx"
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    sheet1_df.to_excel(writer, sheet_name="成績分析表", index=False)
    sheet2_df.to_excel(writer, sheet_name="全班原始成績", index=False)
    sheet3_df.to_excel(writer, sheet_name="boxplot", index=False)
print("Created Successfully!")