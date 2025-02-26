import pandas as pd
import re
from fuzzywuzzy import fuzz, process
import spacy

# 加载 spaCy NER 模型
nlp = spacy.load("en_core_web_sm")


# 使用 NER 提取公司名称中的核心实体部分
def extract_core_name(name):
    doc = nlp(name)
    org_entities = [ent.text for ent in doc.ents if ent.label_ == "ORG"]
    return org_entities[0] if org_entities else name


# 公司名称清理函数
def clean_company_name(name):
    core_name = extract_core_name(name)
    core_name = re.sub(r'\b(Inc|Incorporated|LLC|Ltd|Limited|Corp|Corporation|Co|Company)\b', '', core_name,
                       flags=re.IGNORECASE)
    return core_name.strip().upper()  # 转换为大写


# 模糊匹配函数，加入首字母检查
def fuzzy_match(company_list_1, company_list_2, threshold=80):
    matches = []
    for company1, code1 in company_list_1:
        # 清理公司名称
        clean_name1 = clean_company_name(company1)
        first_letter1 = clean_name1[0] if clean_name1 else ''  # 获取首字母

        # 对列表2的公司名称进行同样的清理
        cleaned_company_list_2 = [(clean_company_name(c[0]), c[1]) for c in company_list_2]

        # 筛选首字母匹配的公司名称
        filtered_list_2 = [c for c in cleaned_company_list_2 if c[0].startswith(first_letter1)]

        # 如果没有匹配的首字母，跳过
        if not filtered_list_2:
            print(f"跳过：'{company1}' 无首字母匹配项。")
            continue

        # 遍历列表2中所有候选项，找出所有得分大于等于阈值的匹配项
        for clean_name2, code2 in filtered_list_2:
            score = fuzz.token_sort_ratio(clean_name1, clean_name2)
            if score >= threshold:
                # 找到一个匹配项，记录下来
                original_name2 = next(c[0] for c in company_list_2 if clean_company_name(c[0]) == clean_name2)
                matches.append({
                    "公司名称1": company1,
                    "代码1": code1,
                    "公司名称2": original_name2,
                    "代码2": code2,
                    "匹配分数": score
                })

                # 打印匹配结果
                print(f"匹配：'{company1}' ({code1}) 与 '{original_name2}' ({code2})，匹配分数：{score}")

    return pd.DataFrame(matches)


# 主处理函数
def main(file1, file2, threshold=90):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)

    # 提取公司名称和代码
    company_list_1 = list(zip(df1['conm'], df1['gvkey']))
    company_list_2 = list(zip(df2['PrivCo_name'], df2['PrivCo_ID']))

    matched_df = fuzzy_match(company_list_1, company_list_2, threshold=threshold)

    result_path = "matched_companies_COMPUSTAT_PrivCo-firstletter.xlsx"
    matched_df.to_excel(result_path, index=False)
    return result_path

# 使用示例
main("COMPUSTAT_NA_Name_ID.xlsx","PrivCo_Name_ID.xlsx", threshold=90)
