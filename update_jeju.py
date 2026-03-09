import pandas as pd

df = pd.read_csv('bidding_winners.csv')
file_name = '입찰결과 - 251113 (종평) 제주외항 2단계(잡화부두) 개발공사_Rev.N4.05.xlsb'
company = '동부건설'
faction = '원숭이(하늘)'

mask = df['File'] == file_name
if mask.any():
    df.loc[mask, 'WinnerCompany'] = company
    df.loc[mask, 'WinnerFaction'] = faction
else:
    new_row = pd.DataFrame([{'File': file_name, 'WinnerCompany': company, 'WinnerFaction': faction}])
    df = pd.concat([df, new_row], ignore_index=True)

df.to_csv('bidding_winners.csv', index=False)
