# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
# pip install tabulate
from tabulate import tabulate

#mps = pd.read_excel('/content/drive/My Drive/MRP.xlsx', sheet_name='MPS', index_col='품목코드') #품목코드를 인덱스로 설정
mps = pd.read_excel('MRP.xlsx', sheet_name='MPS', index_col='품목코드') #품목코드를 인덱스로 설정
mps = mps.sort_index() # 인덱스를 기준으로 정렬

#bom = pd.read_excel('MRP.xlsx', sheet_name='BOM', index_col='Child') #Child 컬럼을 인덱스로 설정
#bom = bom.sort_index() # 인덱스를 기준으로 정렬
bom = pd.read_excel('MRP.xlsx', sheet_name='BOM')

#llc = pd.read_excel('MRP.xlsx', sheet_name='LLC', index_col='품목')
#llc = llc.sort_values('LLC', ascending=True) ## Ture : 오름차순, False : 내림차순
def format_dependencies(df):
    dependencies = {}
    for _, row in df.iterrows():
        child = row['Child']
        parent = row['Parent']
        if child not in dependencies:
            dependencies[child] = set()
        dependencies[child].add(parent)
    return dependencies

def calculate_low_level_codes(df):
    dependencies = format_dependencies(df)
    #print(dependencies)
    low_level_code = {}
    level = 0
    all_items = set(df['Parent']) | set(df['Child'])
    
    while all_items:
        current_level_items = {item for item in all_items if item not in dependencies}
        for item in current_level_items:
            low_level_code[item] = level
        all_items -= current_level_items
        dependencies = {k: v - current_level_items for k, v in dependencies.items() if v - current_level_items}
        level += 1
    
    return low_level_code

low_level_codes = calculate_low_level_codes(bom)

llc = pd.DataFrame(list(low_level_codes.items()), columns=['품목', 'LLC'])
llc = llc.set_index('품목')

bom = bom.set_index('Child')
bom = bom.sort_index()

irf = pd.read_excel('MRP.xlsx', sheet_name='IRF', index_col='품목코드')
irf['기말재고'] = irf['현재재고'] - irf['안전재고']

mrp = pd.read_excel('MRP.xlsx', sheet_name='MRP') # index_col=['품목코드','구분'] 이건 에러가 남.
#mrp = mrp.fillna(method='ffill') # 품목코드 A, B, C, D로 컬럼 다 채우기
mrp = mrp.ffill() # 품목코드 A, B, C, D로 컬럼 다 채우기
mrp = mrp.set_index(['품목코드','구분'])

llc_index = llc.index.tolist()
for item in llc_index:
  #총소요량 계산(생산계획 내용만 업데이트)
  if item in mps.index:
    mps_item = mps.loc[item] # 품목코드가 같은 것만 추출
    for _, row in mps_item.iterrows():
      if(item, '총소요량') in mrp.index and row['예정입고'] in mrp.columns:
        mrp.loc[(item, '총소요량'), row['예정입고']] = row['총소요량']

  #예정입고 계산
  if(irf.loc[(item, '예정입고량')]>0): # 예정입고량이 0 이상 인 경우만 업데이트
    mrp.loc[(item, '예정입고'), irf.loc[(item, '예정입고시기')]] = irf.loc[(item, '예정입고량')]


  for week in range(4,18):
    #총소요량 계산 업데이트(C, D)
    if item in bom.index:
      bom_item = bom.loc[item] # 품목코드가 같은 것만 추출
      for _, row in bom_item.iterrows():
           mrp.loc[(item, '총소요량'), week] += (mrp.loc[(row['Parent'], '계획발주'), week]*row['Qty'])

    if week == 4: # 기말재고를 가지고 계산하는 부분이 필요함.
      #순소요량 계산 = 총소요량 - 전주예상재고 - 예정입고
      mrp.loc[(item, '순소요량'), week] = max(
        mrp.loc[(item, '총소요량'), week]
        - irf.loc[(item, '기말재고')]
        - mrp.loc[(item, '예정입고'), week],
        0)
      #계획수주 계산 =  순소요량, 만약 계획수주량이 최소주문량보다 작을 경우 최소주문량으로 수주
      if mrp.loc[(item, '순소요량'), week] > 0:
        mrp.loc[(item, '계획수주'), week] = max(
              mrp.loc[(item, '순소요량'), week],
              irf.loc[(item, '최소주문량')])

      #예상재고 계산 = 전주예상재고 + 예정입고 + 계획수주 - 총소요량
      mrp.loc[(item, '예상재고'), week] = max(
        irf.loc[(item, '기말재고')]
        + mrp.loc[(item, '예정입고'), week]
        + mrp.loc[(item, '계획수주'), week]
        - mrp.loc[(item, '총소요량'), week],
        0)

    else:
      #순소요량 계산 = 총소요량 - 전주예상재고 - 예정입고
      mrp.loc[(item, '순소요량'), week] = max(
        mrp.loc[(item, '총소요량'), week]
        - mrp.loc[(item, '예상재고'), week-1]
        - mrp.loc[(item, '예정입고'), week],
        0)
      #계획수주 계산 =  순소요량, 만약 계획수주량이 최소주문량보다 작을 경우 최소주문량으로 수주
      if mrp.loc[(item, '순소요량'), week] > 0:
        mrp.loc[(item, '계획수주'), week] = max(
              mrp.loc[(item, '순소요량'), week],
              irf.loc[(item, '최소주문량')])

      #예상재고 계산 = 전주예상재고 + 예정입고 + 계획수주 - 총소요량
      mrp.loc[(item, '예상재고'), week] = max(
        mrp.loc[(item, '예상재고'), week-1]
        + mrp.loc[(item, '예정입고'), week]
        + mrp.loc[(item, '계획수주'), week]
        - mrp.loc[(item, '총소요량'), week],
        0)
      # 계획발주 = 계획수주(week-인도기간)
    if week-irf.loc[(item, '인도기간')] >3: #4주차 부터 표시하기 위해서
      mrp.loc[(item, '계획발주'), week-irf.loc[(item, '인도기간')]] = mrp.loc[(item, '계획수주'), week]
      
print(tabulate(mrp, headers='keys', tablefmt='psql', showindex=True))