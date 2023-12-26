import re


def merge_cells(ws):
    ws.merge_cells('A1:E1')
    ws.merge_cells('F1:H1')
    ws.merge_cells('D2:F2')
    ws.merge_cells('G4:H4')
    ws.merge_cells('G5:H5')
    ws.merge_cells('A6:H6')
    ws.merge_cells('B7:C7')
    ws.merge_cells('G7:H7')
    ws.merge_cells('B8:H8')
    ws.merge_cells('A9:H9')
    ws.merge_cells('B10:C10')
    ws.merge_cells('D10:F10')
    ws.merge_cells('B11:C11')
    ws.merge_cells('D11:F11')
    ws.merge_cells('B12:C12')
    ws.merge_cells('D12:F12')
    ws.merge_cells('B13:C13')
    ws.merge_cells('D13:F13')
    ws.merge_cells('B14:C14')
    ws.merge_cells('D14:F14')
    ws.merge_cells('B15:C15')
    ws.merge_cells('D15:F15')
    ws.merge_cells('B16:C16')
    ws.merge_cells('D16:F16')
    ws.merge_cells('B17:C17')
    ws.merge_cells('D17:F17')
    ws.merge_cells('B18:C18')
    ws.merge_cells('D18:F18')
    ws.merge_cells('B19:C19')
    ws.merge_cells('D19:F19')
    ws.merge_cells('B20:C20')
    ws.merge_cells('D20:F20')
    ws.merge_cells('B21:C21')
    ws.merge_cells('D21:F21')
    ws.merge_cells('B22:C22')
    ws.merge_cells('D22:F22')
    ws.merge_cells('B23:C23')
    ws.merge_cells('D23:F23')
    ws.merge_cells('B24:C24')
    ws.merge_cells('D24:F24')
    ws.merge_cells('B25:C25')
    ws.merge_cells('D25:F25')
    ws.merge_cells('B26:C26')
    ws.merge_cells('D26:F26')
    ws.merge_cells('B27:C27')
    ws.merge_cells('D27:F27')
    ws.merge_cells('B28:C28')
    ws.merge_cells('D28:F28')
    ws.merge_cells('B29:C29')
    ws.merge_cells('D29:F29')
    ws.merge_cells('B30:C30')
    ws.merge_cells('D30:F30')
    ws.merge_cells('B31:C31')
    ws.merge_cells('D31:F31')
    ws.merge_cells('D3:F5')

def set_cells(ws,record):
    ws['A2'] = '设备名称'
    ws['A4'] = '使用部门'
    ws['A7'] = '厂家名称'
    ws['A8'] = '通信地址'
    ws['A9'] = '维修保养记录'
    ws['A10'] = '日期'
    ws['B2'] = '型号'
    ws['B4'] = '启用年月'
    ws['B10'] = '维修 (保养) 原因'
    ws['C2'] = '出厂编号'
    ws['C4'] = '设备原值'
    ws['D2'] = '主要参数'
    ws['D7'] = '电话'
    ws['D10'] = '更换记录'
    ws['F1'] = '建卡时间：'
    ws['F7'] = '传真'
    ws['G2'] = '额定电流'
    ws['G4'] = '设备编号'
    ws['G10'] = '维修人员'
    ws['H2'] = '电机台/功率'
    ws['H10'] = '备注'

    date_pattern = re.compile(r'\b\d{4}-\d{1,2}-\d{1,2}\b')

    #输入内容
    ws['A3'] = record['设备名称']
    ws['A5'] = '麻涌提标'
    ws['A6'] = '制造商联络栏'
    ws['B7'] = record['厂家名称'].replace("nan","")
    # ws['A9'] = record['通信地址']
    # ws['A10'] = record['维修保养记录']
    # ws['A11'] = record['日期']
    ws['B3'] = record['型号']
    ws['B5'] = record['启用年月']
    # ws['B11'] = record['维修 (保养) 原因']
    ws['C3'] = record['出厂编号']
    ws['C5'] = record['设备原值']
    main_ref = record['主要参数']
    main_ref =  '\n'.join([main_ref[i:i+6] for i in range(0, len(main_ref), 6)])
    ws['D3'] = main_ref
    # ws['D8'] = record['电话']
    # ws['D11'] = record['更换记录']
    # ws['F2'] = record['建卡时间：']
    # ws['F8'] = record['传真']
    # ws['G3'] = record['额定电流']
    ws['G5'] = record['设备编号']
    # ws['G11'] = record['维修人员']
    ws['H3'] = record['电机台/功率']
    # ws['H11'] = record['备注']

