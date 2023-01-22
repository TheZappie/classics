import pandas as pd
import itertools as it

SHEET_NAME = 'Uitslagen'

TOEGEVOEGD_DOOR = 'Toegevoegd door'

SONG_NAME = 'Song Name'

p = r"Classics verkiezingen 2022 correctie.xlsx"
dic = {'Azzam': 'Don Santosa',
       'Victor': 'Victor',
       'Jelle': 'Anoniempje',
       'Timo': 'Timo',
       'Tijmen': 'TimmaDoo',
       'Long': 'Long (correct)',
       'Kjeld': ' Kjeld ',
       'Jurriaan': 'Jur',
       'Jochem': 'Ham',
       }

dic2 = {'Ja': True,
        'Nee': False,
        'Weet ik niet': True}

# dtype = {i: 'category' for i in it.chain(dic.values(), ['Toegevoegd door'])}
dtype = {i: 'category' for i in it.chain([TOEGEVOEGD_DOOR])}
converters = {i: lambda x: dic2[x] for i in dic.values()}
df = pd.read_excel(p, dtype=dtype, converters=converters)
df.rename(columns={v: i for i, v in dic.items()}, inplace=True)
df.columns = [SONG_NAME] + df.columns[1:].tolist()
ARTIEST = 'Artiest'
df[[ARTIEST, SONG_NAME]] = df[SONG_NAME].str.strip('[]').str.split(' - ', 1, expand=True)
artists = df.set_index(SONG_NAME)[ARTIEST]
df = df[[SONG_NAME, TOEGEVOEGD_DOOR] + list(dic.keys())]


def func(arg):
    votes = arg[arg.name]
    portion = votes.sum() / len(arg)
    print(f'{arg.name}: {round(portion * 100)} %')


print("Percentage op je eigen nummers gestemd")

df.groupby(TOEGEVOEGD_DOOR, group_keys=True).apply(func)

print("Aantal nummers toegevoegd:")
number_of_songs_added = df[TOEGEVOEGD_DOOR].value_counts()
print(number_of_songs_added.head())

total_votes = df.set_index(SONG_NAME)[dic.keys()].sum(axis=1)
toegevoegd_door = df.set_index(SONG_NAME)[TOEGEVOEGD_DOOR]
total_votes.groupby(toegevoegd_door).mean()

score = total_votes.groupby(toegevoegd_door).mean().round(1).sort_values(ascending=False)
# pd.concat([number_of_songs_added,score], axis=1)

number_of_instant_classics = (total_votes >= 7).groupby(toegevoegd_door).sum()

def f(x):
    if isinstance(x, str):
        return x
    else:
        return x[0]
favorite_artisten = artists.groupby(toegevoegd_door).agg(pd.Series.mode)
favorite_artisten = favorite_artisten.map(f)

result = pd.DataFrame({"Aantal nummers ingebracht": number_of_songs_added,
                       "Score": score,
                       "Aantal instant classics": number_of_instant_classics,
                       "Favorite Artiest": favorite_artisten}
                      ).sort_values("Aantal instant classics", ascending=False)
result.index.name = 'Naam'

writer = pd.ExcelWriter('Results.xlsx', engine='xlsxwriter', )
result.to_excel(writer, sheet_name=SHEET_NAME)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets[SHEET_NAME]

# Get the dimensions of the dataframe.
(max_row, max_col) = result.shape

# Apply a conditional format to the required cell range.
for i in range(1, max_col+1):
    worksheet.conditional_format(1, i, max_row, i,
                                 {'type': '3_color_scale'})

# Close the Pandas Excel writer and output the Excel file.
writer.close()