"""Functions vendored from TeleAI's TableReasoner (SemEval-2025 Task 8, 1st place).

Source: https://github.com/ccx06/TableReasoner (actions/table_desc.py,
actions/entity_associate.py), by Sishi Xiong et al., China Telecom TeleAI.
The repository has no licence file; reuse is by permission of the authors,
obtained by the user on 2026-11-03.

Copied as published except where a line is marked "change:":
- tables come in as a DataFrame instead of being re-read from a file path;
- random sampling is seeded, so runs are reproducible (competition Rule 8);
- a trailing comma that wrapped `cell_example` in a tuple is removed;
- the candidate filter and example-row sizing are lifted out of MetaGPT
  Action classes into plain functions.
"""

import random

import numpy as np
import pandas as pd

SEED = 0  # change: fixed seed for all sampling


def retrive_column_feature(df, columns_list, category_keep_num=15):
    new_df = []
    for column in columns_list:
        col_type = df[column].dtype
        example = None

        # for boolean data
        if pd.api.types.is_bool_dtype(col_type):
            counter = df[column].value_counts()
            example = {
                "True Num": counter[True].item() if True in counter else 0,
                "False Num": counter[False].item() if False in counter else 0
            }

        # for numerical data
        elif pd.api.types.is_numeric_dtype(col_type):
            example = {
                "minimum_value": df[column].min().item() if isinstance(df[column].min(), np.int64) else df[column].min(),
                "maximum_value": df[column].max().item() if isinstance(df[column].min(), np.int64) else df[column].max(),
                "median_value": df[column].median().item() if isinstance(df[column].min(), np.int64) else df[column].median(),
                "average_value": df[column].mean().item() if isinstance(df[column].min(), np.int64) else df[column].mean(),
            }

        # for date/time data
        elif pd.api.types.is_datetime64_any_dtype(col_type):
            example = {
                "earliest_date": df[column].min().strftime('%Y-%m-%d %H:%M:%S'),
                "latest_date": df[column].max().strftime('%Y-%m-%d %H:%M:%S')
            }

        # for categorical data
        elif pd.api.types.is_object_dtype(col_type):
            value_counts = df[column].value_counts()
            if len(value_counts) <= category_keep_num:
                categories = value_counts.index.tolist()
                example = {
                    "all_categories": categories
                }
            else:
                top_categories = value_counts.nlargest(category_keep_num).index.tolist()
                example = {
                    "category_examples": top_categories,
                    "total_number_of_categories": len(value_counts),     # ?
                    "caption": f"There are many types in this column, displaying the top {category_keep_num} most common types"
                }

        new_df.append({
            "column_name": column,
            "dtype": str(col_type),
            "example": example
        })

    result_df = pd.DataFrame(new_df)
    return result_df


def get_table_schema(df):
    columns_list = df.columns.tolist()
    if len(columns_list) <= 20:
        result_df = retrive_column_feature(df, columns_list, category_keep_num=10)
    elif len(columns_list) <= 50:
        result_df = retrive_column_feature(df, columns_list, category_keep_num=6)
    else:
        result_df = retrive_column_feature(df, columns_list, category_keep_num=4)

    return result_df.to_dict('records')


def get_refined_table_schema(table_schema, relevant_column_list, df):  # change: df passed in
    """ Simplify the table based on relevant columns and obtain more detailed information for relevant columns. """
    rng = random.Random(SEED)  # change: seeded
    refined_table_schema = {
        "file_path": table_schema['file_path'],
        "table_name": table_schema.get('table_name', ''),
        "table_description": table_schema.get('table_description', ''),
        "number_of_rows": table_schema.get('number_of_rows', ''),
        "column_list": [],
        "cell_example": [],
        "column_description": [],
    }

    # step 1: Read information about relevant columns from table schema.
    try:
        column_list = table_schema['column_list']
        for column_name in relevant_column_list:
            if column_name in column_list:
                refined_table_schema['column_list'].append(column_name)
            else:
                print(f'Wrong column name!! {column_name} not in the table.')
        for column_desc in table_schema['column_description']:
            if column_desc['column_name'] in relevant_column_list:
                refined_table_schema['column_description'].append(column_desc)

        # step 2: Update detailed information about the relevant columns -- column_description
        new_column_description = refined_table_schema['column_description']
        for i, each_column_desc in enumerate(refined_table_schema['column_description']):
            if each_column_desc['dtype'] != 'object':
                continue
            column = each_column_desc['column_name']
            value_counts = df[column].value_counts()
            category_max_count = 50  # 100
            if len(str(value_counts.index.tolist()[0]).split(' ')) > 50 \
                or len(str(value_counts.index.tolist()[rng.sample(range(len(value_counts)), k=1)[0]]).split(' ')) > 25:   # 50
                category_max_count = 20
            elif len(str(value_counts.index.tolist()[0]).split(' ')) > 20 \
                or len(str(value_counts.index.tolist()[rng.sample(range(len(value_counts)), k=1)[0]]).split(' ')) > 10:   # 20
                category_max_count = 50

            if len(value_counts) <= category_max_count:
                categories = value_counts.index.tolist()
                new_column_description[i]['all_categories'] = categories
                new_column_description[i]['total_number_of_categories'] = len(categories)
                if 'category_examples' in new_column_description[i]:
                    del new_column_description[i]['category_examples']
                if 'example' in new_column_description[i]:
                    del new_column_description[i]['example']

            else:
                top_categories = value_counts.nlargest(category_max_count).index.tolist()
                new_column_description[i]['category_examples'] = top_categories
                new_column_description[i]['total_number_of_categories'] = len(value_counts)
                new_column_description[i]['caption'] = f"There are many types in this column, displaying the top {category_max_count} most common types"
                if 'all_categories' in new_column_description[i]:
                    del new_column_description[i]['all_categories']
                if 'example' in new_column_description[i]:
                    del new_column_description[i]['example']

            refined_table_schema['column_description'] = new_column_description
        # change: seeded sample, trailing tuple comma removed, set once after the loop
        # (was inside it, so tables without text columns got no example rows)
        refined_table_schema['cell_example'] = df.sample(10, random_state=SEED)[refined_table_schema['column_list']].to_dict('records') if len(df)>20 else df[:][refined_table_schema['column_list']].to_dict('records')

    except Exception as e:
        print(f'Table Schema Refine Error!!\n {e}')
        columns_list = [c for c in df.columns.tolist() if c in relevant_column_list]
        # Update column_list, cell_example
        refined_table_schema['column_list'] = columns_list
        refined_table_schema['cell_example'] = df.sample(10, random_state=SEED)[refined_table_schema['column_list']].to_dict('records') if len(df)>20 else df[:][refined_table_schema['column_list']].to_dict('records')

        result_df = retrive_column_feature(df, columns_list, category_keep_num=100)
        for idx, c in result_df.iterrows():
            refined_table_schema['column_description'].append({
                'column_name': c['column_name'],
                'dtype': c['dtype']
            })
            for k, v in c['example'].items():
                refined_table_schema['column_description'][idx][k] = v

    return refined_table_schema


def cell_example_count(df):
    """change: lifted from TableDesc.run -- how many random example rows to show."""
    column_num = len(df.columns.tolist())
    cell_example_len = sum([len(str(v).split()) for v in df.sample(1, random_state=SEED).to_dict('records')[0].values()])
    if column_num <= 20 and cell_example_len <= 300:
        cell_chosen_num = 10
    elif column_num <= 50 and cell_example_len <= 500:
        cell_chosen_num = 5
    else:
        cell_chosen_num = 2
    return cell_chosen_num


def longest_common_subsequence(text1: str, text2: str) -> int:
    m, n = len(text1), len(text2)
    dp = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if text1[i - 1] == text2[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    return dp[m][n]


def entity_candidates(entity, column_ele_list):
    """change: lifted from EntityLinking.run -- values of one column an entity may refer to."""
    candidates = []
    text1 = entity.lower().replace(' ', '')
    for ele in column_ele_list:
        text2 = ele.lower().replace(' ', '')
        long_com = longest_common_subsequence(text1, text2)
        if long_com > 1 and long_com/len(text1) > 0.6:
            candidates.append((ele, long_com))
    candidates.sort(key=lambda x:x[1], reverse=True)
    candidates = candidates[:10]
    return candidates
