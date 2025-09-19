import pandas as pd

def extract_sec_mapping(excel_file, sheet_name="Sec Mapping"):
    try:
        # 读取XLSM文件的指定工作表，保持数据原样
        df = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            engine="openpyxl",  # 明确指定引擎以支持xlsm格式
            dtype=str,          # 所有列都按字符串处理，避免数值格式转换问题
            keep_default_na=False  # 不将空值转换为NaN，保持为空字符串
        )
        
        # 确保至少有两列数据
        if len(df.columns) < 2:
            raise ValueError("工作表至少需要包含两列数据")
        
        # 获取前两列的列名
        col1, col2 = df.columns[0], df.columns[1]
        
        # 转换为字典
        sec_mapping = {}
        for _, row in df.iterrows():
            key = row[col1].strip()  # 去除首尾空格
            value = row[col2].strip()
            sec_mapping[key] = value
        
        # 将字典写入seg_mapping_config.py文件
        with open("seg_mapping_config.py", "w", encoding="utf-8") as f:
            f.write("# 从VACB_FIS Recon 20250901.xlsm提取的Security映射关系\n")
            f.write("# 第一列: %s, 第二列: %s\n" % (col1, col2))
            f.write("sec_mapping = {\n")
            for key, value in sec_mapping.items():
                # 处理字符串中的引号和特殊字符
                key_str = key.replace('"', '\\"').replace('\n', '\\n')
                value_str = value.replace('"', '\\"').replace('\n', '\\n')
                f.write(f'    "{key_str}": "{value_str}",\n')
            f.write("}\n")
        
        print(f"成功生成配置文件: seg_mapping_config.py，共包含 {len(sec_mapping)} 条记录")
        
    except FileNotFoundError:
        print(f"错误: 未找到文件 {excel_file}")
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")

if __name__ == "__main__":
    # Excel文件路径（.xlsm格式）
    excel_file_path = "VACB_FIS Recon 20250901.xlsm"
    extract_sec_mapping(excel_file_path)
