import streamlit as st
import pandas as pd
import io
from difflib import SequenceMatcher

def get_best_matches_with_code(target, choices_dict, top_n=3):
    """
    Return the top_n best matching strings with their corresponding codes.
    """
    match_ratios = [(choice, SequenceMatcher(None, target, choice).ratio(), choices_dict[choice]) 
                    for choice in choices_dict.keys()]
    match_ratios.sort(key=lambda x: x[1], reverse=True)
    return [(match[0], match[2]) for match in match_ratios[:top_n]]

def filter_incorrect_matches(df):
    """
    篩選出不正確的比對結果
    """
    condition = (df['行業比對結果'] != '資料正確') | (df['職業比對結果'] != '資料正確')
    filtered_df = df[condition].copy()
    return filtered_df

def process_comparison(reference_data, uploaded_data):
    """
    Process the comparison between reference and uploaded data.
    """
    # Create dictionaries for quick lookup
    original_jobs_dict = reference_data.set_index('職業')['職註'].to_dict()
    original_industries_dict = reference_data.set_index('行業')['行註'].to_dict()
    
    # Process data in chunks
    chunk_size = 500
    chunks = [uploaded_data[i:i + chunk_size] for i in range(0, uploaded_data.shape[0], chunk_size)]
    
    all_job_results = []
    all_industry_results = []
    
    for chunk in chunks:
        # Process jobs
        job_results = []
        for _, row in chunk.iterrows():
            if row['職業'] in original_jobs_dict and original_jobs_dict[row['職業']] == row['職註']:
                job_results.append('資料正確')
            else:
                best_matches = get_best_matches_with_code(row['職業'], original_jobs_dict)
                job_results.append(best_matches)
        
        # Process industries
        industry_results = []
        for _, row in chunk.iterrows():
            if row['行業'] in original_industries_dict and original_industries_dict[row['行業']] == row['行註']:
                industry_results.append('資料正確')
            else:
                best_matches = get_best_matches_with_code(row['行業'], original_industries_dict)
                industry_results.append(best_matches)
        
        all_job_results.extend(job_results)
        all_industry_results.extend(industry_results)
    
    uploaded_data['行業比對結果'] = all_industry_results
    uploaded_data['職業比對結果'] = all_job_results
    
    # 進行第二次篩選
    filtered_results = filter_incorrect_matches(uploaded_data)
    
    return uploaded_data, filtered_results

def main():
    st.title('行業職業代碼比對系統')
    
    # 載入參考資料（直接從專案目錄讀取）
    try:
        reference_data = pd.read_excel('reference_data/indu_occu_search.xlsx')
        st.success('參考資料載入成功！')
    except Exception as e:
        st.error(f'載入參考資料時發生錯誤: {str(e)}')
        return
    
    # 檔案上傳介面
    uploaded_file = st.file_uploader('請上傳要比對的Excel檔案', type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            # 讀取上傳的檔案
            uploaded_data = pd.read_excel(uploaded_file)
            
            # 執行比對
            full_results, filtered_results = process_comparison(reference_data, uploaded_data)
            
            # 顯示比對結果預覽
            st.write('完整比對結果預覽：')
            st.dataframe(full_results.head())
            
            st.write('篩選後的比對結果預覽（僅顯示不正確的項目）：')
            st.dataframe(filtered_results.head())
            
            # 準備下載檔案
            output_full = io.BytesIO()
            full_results.to_excel(output_full, index=False)
            output_full.seek(0)
            
            output_filtered = io.BytesIO()
            filtered_results.to_excel(output_filtered, index=False)
            output_filtered.seek(0)
            
            # 下載按鈕
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label='下載完整比對結果',
                    data=output_full,
                    file_name='full_comparison_results.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            
            with col2:
                st.download_button(
                    label='下載篩選後的比對結果',
                    data=output_filtered,
                    file_name='filtered_comparison_results.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            
        except Exception as e:
            st.error(f'處理過程發生錯誤: {str(e)}')

if __name__ == '__main__':
    main()
