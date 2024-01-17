from pathlib import Path

import pandas as pd
import itertools as it

THRESHOLD_2023 = 5

ARTIESTEN_RESULTATEN = 'Artiesten_Resultaten'
SHEET_NAME = 'Uitslagen'
TOEGEVOEGD_DOOR = 'Added By'
SONG_NAME = 'Track Name'
ARTIEST = 'Artist Name(s)'

data_path = r"data/2023.xlsx"
path_results = 'Results2023'

naming_2023 = {'Azzam': 'Don Santosa',
               'Victor': 'Victor',
               'Jelle': 'Haicolientje',
               'Timo': 'Timo',
               'Tijmen': 'Tijmen "TimmaDoo" Post',
               'Long': 'Long',
               # 'Kjeld': ' Kjeld',
               'Jurriaan': 'Navelpluis',
               'Jochem': 'Ham',
               }

n = len(naming_2023.keys())

dic2 = {'Ja': True,
        'Nee': False,
        'Weet ik niet': True}

dtype = {i: 'category' for i in it.chain([TOEGEVOEGD_DOOR])}
converters = {i: lambda x: dic2[x] for i in naming_2023.values()}
df = pd.read_excel(data_path, dtype=dtype, converters=converters, index_col=0)
df.rename(columns={v: i for i, v in naming_2023.items()}, inplace=True)

# df[[ARTIEST, SONG_NAME]] = df[SONG_NAME].str.strip('[]').str.split(' - ', n=1, expand=True)
artists = df.set_index(SONG_NAME)[ARTIEST]
df = df[[SONG_NAME, TOEGEVOEGD_DOOR] + list(naming_2023.keys())]


def nominated_by(dataframe: pd.DataFrame):
    if not dataframe.name in dataframe.columns:
        return
    votes = dataframe[dataframe.name]
    portion = votes.sum() / len(dataframe)
    print(f'{dataframe.name}: {round(portion * 100)} %')


def consistency(dataframe: pd.DataFrame):
    if not dataframe.name in dataframe.columns:
        return
    votes = dataframe[dataframe.name]
    portion = votes.sum() / len(dataframe)
    return (f'{round(portion * 100)} %')


df.groupby(TOEGEVOEGD_DOOR, group_keys=True, observed=False).apply(nominated_by)

consistentie = df.groupby(TOEGEVOEGD_DOOR, group_keys=True, observed=False).apply(consistency)
number_of_songs_added = df[TOEGEVOEGD_DOOR].value_counts()
total_votes = df.set_index(SONG_NAME)[naming_2023.keys()].sum(axis=1)
toegevoegd_door = df.set_index(SONG_NAME)[TOEGEVOEGD_DOOR]
total_votes.groupby(toegevoegd_door, observed=False).mean()

score = total_votes.groupby(toegevoegd_door, observed=False).mean().round(1).sort_values(ascending=False)

number_of_instant_classics = (total_votes >= THRESHOLD_2023).groupby(toegevoegd_door, observed=False).sum()


def f(x):
    if isinstance(x, str):
        return x
    else:
        return x[0]


favorite_artisten = artists.groupby(toegevoegd_door, observed=False).agg(pd.Series.mode)
favorite_artisten = favorite_artisten.map(f)

result = pd.DataFrame({"Aantal instant classics": number_of_instant_classics,
                       "Aantal nummers ingebracht": number_of_songs_added,
                       f"Gemiddelde rating (0-{n})": score,
                       "Favorite Artiest": favorite_artisten,
                       "Consistent": consistentie}
                      ).sort_values("Aantal instant classics", ascending=False)
result.index.name = 'Naam'

artist_inbrengen = artists.value_counts()
artist_rating = total_votes.groupby(artists).mean()
artist_results = pd.DataFrame({"Aantal nummers ingebracht": artist_inbrengen,
                               f"Gemiddelde rating (0-{n})": artist_rating},
                              )
artist_results = artist_results.sort_values(f"Aantal nummers ingebracht", ascending=False)

path_results = Path(path_results)
result.to_markdown(path_results.with_suffix('.md'))
writer = pd.ExcelWriter(path_results.with_suffix('.xlsx'), engine='xlsxwriter', )
result.to_excel(writer, sheet_name=SHEET_NAME)
workbook = writer.book
worksheet = writer.sheets[SHEET_NAME]

(max_row, max_col) = result.shape
for i in range(1, max_col + 1):
    worksheet.conditional_format(1, i, max_row, i,
                                 {'type': '3_color_scale'})

artist_results.to_excel(writer, sheet_name=ARTIESTEN_RESULTATEN)
worksheet = writer.sheets[ARTIESTEN_RESULTATEN]
(max_row, max_col) = artist_results.shape
for i in range(1, max_col + 1):
    worksheet.conditional_format(1, i, max_row, i,
                                 {'type': '3_color_scale'})
writer.close()
