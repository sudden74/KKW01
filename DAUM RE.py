'''
크롤링 데이터 가공과 결측치 처리 예제 (Daum 부동산 예제)
Q: DataFrame의 두 행(row)이 하나의 데이터를 이룹니다. 짝수 행의 데이터를 추출하여 컬럼으로 만드려면 어떻게 하나요?
A: DataFrame.ix[] 인덱싱을 사용하여 짝수행 전체를 얻어 가공하고, pd.concat()을 이용여 합칠 수 있습니다.
'''

import pandas as pd

url = "http://realestate.daum.net/iframe/maemul/maemulList.daum?areaCode=2130010&mcateCode=A1&saleTypeCode=S&tabName=maemulList&isExpanded=false&page=1"
dfs = pd.read_html(url)

# 2개의 DataFrame을 읽었다
len(dfs)

'''
두번째 DataFrame을 얻는다.
총 60 행(row)이지만, 2개의 행(row)가 하나의 데이터를 이루고 있다. 짝수번째 행(row)는 부가적인 설명이 있다.
'''
df = dfs[1]
print(len(df), 'rows')

df.head(10)

'''
짝수번째 행(인덱스가 1,3,5,7..)은 NaN을 포함하고 있다.
가장 간단한 방법은 NaN을 가지고 있는 행을 제거(drop)하는 것이다. (DataFrame.dropna)
'''
df_tmp = df.dropna()
print(len(df_tmp), 'rows')

df_tmp.head(10)

'''
DataFrame의 행,열 조작
'''
# ix[] 인덱싱을 사용하여 짝수번째 행(row)만 추출한다.
df1 = df.ix[::2,:]

print(len(df1), 'rows')
df1.head(10)

# 이번에는 홀수번째 행(row)만 추출한다.
df2 = df.ix[1::2, 0]
df2.name = '비고'
df2.head(10)

# df1과 df2를 합치기(concat)위해 위해 index를 동일하게 초기화 한다.
df1.reset_index(drop=True , inplace=True)
df2.reset_index(drop=True , inplace=True)

# df1 과  df2 를 컬럼으로 합치기
result = pd.concat([df1, df2], axis=1)
result.head(10)

'''
컬럼 다듬기
'최초게재일' 컬럼에 거래종류와 날짜가 함께 있다. 거래종류는 매매, 전세, 월세 등이 있다. 이 데이터를 둘로 나누어 각각 '거래종류'를 추출하고, '최초게제일'에는 날짜만 남도록 해보자.
Series.str.extract() 에 정규식을 사용하여 매치되는 문자열을 추출할 수 있다.
'''

# '매매 17.03.01' 데이터에 대해,
# 정규식 '(.*) '을 적용하면 '매매'만 추출
result['거래종류'] = result['최초게재일'].str.extract('(.*) ')
result.head(10)

# '매매 17.03.01' 데이터에 대해,
# 정규식 '(\d+.\d+.\d+)'을 적용하여 추출하면 '17.03.01'만 추출
result['최초게재일'] = result['최초게재일'].str.extract('(\d+.\d+.\d+)')
result.head(10)

