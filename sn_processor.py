import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from typing import List, Dict, Callable
import sqlite3
import os

class SNProcessor:
    def __init__(self, progress_callback: Callable[[float], None] = None, 
                 status_callback: Callable[[str], None] = None):
        self.db_conn = sqlite3.connect('sn_database.db')
        self.progress_callback = progress_callback
        self.status_callback = status_callback
        self.setup_databases()
        
    def setup_databases(self):
        """Initialize the required database tables"""
        # Table for storing SN code groups
        self.db_conn.execute('''
            CREATE TABLE IF NOT EXISTS sn_groups (
                group_id INTEGER PRIMARY KEY,
                sn_codes TEXT
            )
        ''')
        
        # Table for storing query results
        self.db_conn.execute('''
            CREATE TABLE IF NOT EXISTS query_results (
                id INTEGER PRIMARY KEY,
                shipping_barcode TEXT,
                model_number TEXT,
                description TEXT,
                service_start_time TEXT,
                service_end_time TEXT,
                service_package_name TEXT,
                cocare_service_csp TEXT
            )
        ''')
        self.db_conn.commit()

    def split_sn_codes(self, sn_codes: List[str]) -> Dict[int, List[str]]:
        """Split SN codes into groups of 20"""
        groups = {}
        for i in range(0, len(sn_codes), 20):
            group_id = (i // 20) + 1
            groups[group_id] = sn_codes[i:i+20]
        return groups

    def store_sn_groups(self, groups: Dict[int, List[str]]):
        """Store SN code groups in database"""
        for group_id, sn_codes in groups.items():
            self.db_conn.execute(
                'INSERT INTO sn_groups (group_id, sn_codes) VALUES (?, ?)',
                (group_id, ','.join(sn_codes))
            )
        self.db_conn.commit()

    def query_huawei_support(self, sn_codes: List[str]) -> List[Dict]:
        """Query Huawei support website for SN codes"""
        if self.status_callback:
            self.status_callback("正在启动浏览器...")
            
        driver = webdriver.Chrome()
        try:
            driver.get('https://support.huawei.com/enterprise/ecareWechat')
            
            if self.status_callback:
                self.status_callback("正在输入SN码...")
                
            # Wait for input field and enter SN codes
            sn_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'borderNum'))
            )
            sn_input.send_keys('\n'.join(sn_codes))
            
            if self.status_callback:
                self.status_callback("等待验证码输入...")
                
            # Handle captcha (requires manual intervention or OCR integration)
            time.sleep(10)  # Adjust time as needed for manual captcha entry
            
            if self.status_callback:
                self.status_callback("正在查询...")
                
            # Submit query
            submit_button = driver.find_element(By.ID, 'btnSearch')
            submit_button.click()
            
            # Wait for results and extract data
            results_table = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'IGridTable'))
            )
            
            if self.status_callback:
                self.status_callback("正在提取数据...")
                
            # Extract and process results
            results = []
            rows = results_table.find_elements(By.TAG_NAME, 'tr')[1:]  # Skip header
            for row in rows:
                cols = row.find_elements(By.TAG_NAME, 'td')
                result = {
                    'shipping_barcode': cols[0].text,
                    'model_number': cols[1].text,
                    'description': cols[2].text,
                    'service_start_time': cols[3].text,
                    'service_end_time': cols[4].text,
                    'service_package_name': cols[5].text,
                    'cocare_service_csp': cols[6].text
                }
                results.append(result)
            
            return results
            
        finally:
            driver.quit()

    def store_results(self, results: List[Dict]):
        """Store query results in database"""
        if self.status_callback:
            self.status_callback("正在保存结果...")
            
        for result in results:
            self.db_conn.execute('''
                INSERT INTO query_results 
                (shipping_barcode, model_number, description, 
                service_start_time, service_end_time, 
                service_package_name, cocare_service_csp)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (
                result['shipping_barcode'], result['model_number'],
                result['description'], result['service_start_time'],
                result['service_end_time'], result['service_package_name'],
                result['cocare_service_csp']
            ))
        self.db_conn.commit()

    def export_results(self, filename: str = 'query_results.xlsx'):
        """Export results to Excel file"""
        if self.status_callback:
            self.status_callback("正在导出结果...")
            
        query = '''
            SELECT * FROM query_results
        '''
        df = pd.read_sql_query(query, self.db_conn)
        df.to_excel(filename, index=False)

    def process_sn_codes(self, sn_codes: List[str]):
        """Main process to handle SN code processing"""
        # Split and store SN codes
        groups = self.split_sn_codes(sn_codes)
        total_groups = len(groups)
        self.store_sn_groups(groups)
        
        # Process each group
        for i, (group_id, group_codes) in enumerate(groups.items(), 1):
            if self.status_callback:
                self.status_callback(f"正在处理第 {i}/{total_groups} 组...")
            if self.progress_callback:
                self.progress_callback((i - 1) / total_groups * 100)
                
            results = self.query_huawei_support(group_codes)
            self.store_results(results)
            
            if self.progress_callback:
                self.progress_callback(i / total_groups * 100)
        
        # Export final results
        self.export_results()
        
        if self.status_callback:
            self.status_callback("处理完成")