import configparser
import pandas as pd
import os

class RA_Statistic:
    def __init__(self, config):
        self.config = config
        self.excelfile_with_score = config['ExcelFile']['Score']
        self.excelfile_with_admission = config['ExcelFile']['AdmissionChannels']
        self.Score_df = self.__load_score()
        self.Admission_df = self.__load_admission()
    
    def __load_score(self):
        # load dataframe
        Score_df = pd.read_excel(f'./data/{self.excelfile_with_score}.xls', header=None)
        # Set Score columns name
        Score_df.columns = Score_df.iloc[2]
        Score_df = Score_df.drop(index=[0,1,2,3]).reset_index(drop=True) # Delete
        # Filter 停休、休學...
        Score_df = Score_df[["學號", "姓名", "總成績"]]  
        Score_df = Score_df.dropna(subset=["總成績"])  # remove NaN
        Score_df["總成績"] = Score_df["總成績"].astype(int)
        return Score_df
    def __load_admission(self):
        # load dataframe
        Admission_df = pd.read_excel(f'./data/{self.excelfile_with_admission}.xls', header=None)
        # Set Score columns name
        Admission_df.columns = Admission_df.iloc[7]
        Admission_df = Admission_df.drop(index=list(range(8))).reset_index(drop=True) # Delete
        Admission_df = Admission_df[["學號", "入學管道"]]
        return Admission_df

    def statistics(self):
        merged_df = pd.merge(self.Score_df, self.Admission_df, on="學號", how="inner")
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
        # print(merged_df.to_string())
        self.merged_df = merged_df
        self.sheet1()
        self.sheet2()
        self.sheet3()



    def sheet1(self):
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

        self.sheet1_df = calculate_admission_statistics(self.merged_df)


    def sheet2(self):
        sheet2_df = self.merged_df[["學期","課號","班別","課名","入學管道","總成績","名次","及格","前30%"]]
        sheet2_df = sheet2_df.rename(columns={"總成績":"分數"})
        self.sheet2_df = sheet2_df

    def sheet3(self):
        additional_data = self.sheet2_df.copy()
        additional_data["入學管道"] = "全班"
        # merge oringal df and addional df
        sheet3_df = pd.concat([self.sheet2_df, additional_data], ignore_index=True)
        sheet3_df = sheet3_df[["學期", "課號", "班別", "課名", "入學管道", "分數"]]
        self.sheet3_df = sheet3_df

    def saveFile(self):
        output_file = f"{config['ClassInfo']['Semester']}-{config['ClassInfo']['ClassID']}{config['ClassInfo']['ClassType']}{config['ClassInfo']['ClassName']}.xlsx"
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            self.sheet1_df.to_excel(writer, sheet_name="成績分析表", index=False)
            self.sheet2_df.to_excel(writer, sheet_name="全班原始成績", index=False)
            self.sheet3_df.to_excel(writer, sheet_name="boxplot", index=False)
        print("Created Successfully!")

if __name__ == '__main__':
    os.makedirs('data', exist_ok=True)
    config_file = "config.ini"
    if not os.path.exists(config_file):
        config = configparser.ConfigParser()
        config["ExcelFile"] = {
            "Score": "(成績下載)CSxxxA_113_1_score",
            "AdmissionChannels": "(入學管道)1131_CSxxx_A"
        }

        config["ClassInfo"] = {
            "Semester": "1131",
            "ClassID": "CSxxx",
            "ClassType": "A",
            "ClassName": "(課程名稱)xxxx"
        }

        with open(config_file, "w", encoding="utf-8") as file:
            config.write(file)
            print("Input class information")
            exit(0)

    config = configparser.ConfigParser()
    config.read(config_file, encoding="utf-8")
    RA = RA_Statistic(config)
    RA.statistics()
    RA.saveFile()
    