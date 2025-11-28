#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# Bỏ tất cả warnings
import warnings
warnings.filterwarnings('ignore')

from sql_helpers_new import get_cd_details_df, PrintCD_TKN, PrintCD_TKX

def main():
    # Thông tin kết nối
    Sqlhost = ("192.168.100.6,1433", "Ecus5vnaccs_liem", "sa1", "12345678sa")
    
    # ID tờ khai cần test
    dtokhaimdid = 40109
    
    try:
        # Lấy dữ liệu từ database
        data = get_cd_details_df(Sqlhost, dtokhaimdid)
        
        # Kiểm tra _XorN để quyết định gọi hàm nào
        if not data['dtokhaimd'].empty:
            dtokhaimd_row = data['dtokhaimd'].iloc[0]
            
            if '_XorN' in dtokhaimd_row:
                xor_n_value = dtokhaimd_row['_XorN']
                
                if xor_n_value == 'X':
                    PrintCD_TKX(dtokhaimdid, 1, data, 'output')
                elif xor_n_value == 'N':
                    PrintCD_TKN(dtokhaimdid, 1, data, 'output')
                else:
                    print(f"→ Giá trị _XorN không hợp lệ: {xor_n_value}")
            else:
                print("→ Không tìm thấy trường _XorN trong dữ liệu")
        else:
            print("→ Không có dữ liệu từ bảng dtokhaimd")
            
    except Exception as e:
        print(f"✗ Lỗi: {e}")
    
    print("✓ Test hoàn thành!")

if __name__ == "__main__":
    main() 