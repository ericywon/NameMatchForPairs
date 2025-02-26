import pandas as pd
import re
from fuzzywuzzy import fuzz, process
import spacy

# 加载 spaCy 的英文 NER 模型
nlp = spacy.load("en_core_web_sm")


# 上传和读取 Excel 文件
def read_excel(file_path):
    return pd.read_excel(file_path)


# 使用 NER 提取公司名称中的核心实体部分
def extract_core_name(name):
    doc = nlp(name)
    org_entities = [ent.text for ent in doc.ents if ent.label_ == "ORG"]
    # 若找到“ORG”实体，返回第一个；否则返回原始名称
    return org_entities[0] if org_entities else name


# 公司名称清理函数：去除公司类型并转换为大写
def clean_company_name(name):
    # 使用 NER 提取核心名称部分
    core_name = extract_core_name(name)
    # 去除常见的公司类型词汇并转换为大写
    core_name = re.sub(r'\b(Inc|Incorporated|LLC|Ltd|Limited|Corp|Corporation|Co|Company)\b', '', core_name,
                       flags=re.IGNORECASE)
    return core_name.strip().upper()  # 转换为大写


# 模糊匹配函数，加入公司名称清理步骤
def fuzzy_match(company_list_1, company_list_2, threshold=80):
    matches = []
    for company1, code1 in company_list_1:
        # 清理公司名称并转换为大写
        clean_name1 = clean_company_name(company1)

        # 对列表2的公司名称进行同样的清理
        cleaned_company_list_2 = [(clean_company_name(c[0]), c[1]) for c in company_list_2]

        # 执行模糊匹配
        best_match = process.extractOne(clean_name1, [c[0] for c in cleaned_company_list_2],
                                        scorer=fuzz.token_sort_ratio)
        if best_match and best_match[1] >= threshold:
            company2_index = [c[0] for c in cleaned_company_list_2].index(best_match[0])
            original_company2, code2 = company_list_2[company2_index]

            # 打印实时匹配结果
            print(f"匹配：'{company1}' ({code1}) 与 '{original_company2}' ({code2})，匹配分数：{best_match[1]}")

            matches.append({
                "conm": company1,
                "gvkey": code1,
                "PrivCo_name": original_company2,
                "PrivCo_ID": code2,
                "Matching_Score": best_match[1]
            })
    return pd.DataFrame(matches)


# 主函数：上传文件并进行模糊匹配
def main(file_path_1, file_path_2, threshold=80):
    # 读取文件内容
    df1 = read_excel(file_path_1)
    df2 = read_excel(file_path_2)

    # 提取公司名称和代码
    company_list_1 = list(zip(df1['conm'], df1['gvkey']))
    company_list_2 = list(zip(df2['PrivCo_name'], df2['PrivCo_ID']))

    # 进行模糊匹配
    matched_df = fuzzy_match(company_list_1, company_list_2, threshold=threshold)

    # 保存匹配结果到 Excel
    matched_df.to_excel("matched_companies_COMPUSTAT_PrivCo.xlsx", index=False)
    print("匹配结果已保存到 matched_companies_COMPUSTAT_PrivCo.xlsx")

# 使用示例
main("COMPUSTAT_NA_Name_ID.xlsx", "PrivCo_Name_ID.xlsx", threshold=95)
