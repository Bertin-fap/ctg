def plot_club_38(ctg_path):

    df = pd.read_excel(ctg_path / Path('club_38.xlsx'))
    path_villes_de_france = Path(__file__).parent.parent / Path('ctgfuncts') / Path('CTG_RefFiles/villes_france_premium.csv')
    df_villes = pd.read_csv(path_villes_de_france,header=None,usecols=[2,19,20])
    dic_long = dict(zip(df_villes[2] , df_villes[19]))
    dic_lat = dict(zip(df_villes[2] , df_villes[20]))

    df['Ville'] = df['Ville'].str.replace(' ','-')
    df['Ville'] = df['Ville'].str.replace('ST-','SAINT-')
    df['Ville'] = df['Ville'].str.replace('\-D\-+',"-D'",regex=True)
    df['Ville'] = df['Ville'].str.replace('^LA-',"LA ",regex=True)
    df['Ville'] = df['Ville'].str.replace('^LE-',"LE ",regex=True)
    df['Ville'] = df['Ville'].str.replace('SAINT-HILAIRE-DU-TOUVET',"SAINT-HILAIRE",regex=False)
    df['Ville'] = df['Ville'].str.replace('SAINT-HILAIRE',"SAINT-HILAIRE-38",regex=False)
    df['Ville'] = df['Ville'].str.replace('LAVAL',"LAVAL-38",regex=False)
    df['Ville'] = df['Ville'].str.replace('LES-ABRETS',"LES ABRETS",regex=False)
    df['Ville'] = df['Ville'].str.lower()

    df['long'] = df['Ville'].map(dic_long)
    df['lat'] = df['Ville'].map(dic_lat)



    kol = folium.Map(location=[45.2,5.7], tiles='openstreetmap', zoom_start=12)

    for latitude,longitude,size, ville, num_ffct, club in zip(df['lat'],
                                                        df['long'],
                                                        df['number'],
                                                        df['Ville'],
                                                        df['N° FFCT'],
                                                        df['Nom Club'] ):

        long_ville, lat_ville =df.query("Ville==@ville")[['long','lat']].values[0]#.flatten()
        color='blue'

        folium.Circle(
                        location=[latitude, longitude],
                        radius=size*10,
                        popup=f'{ville} ({size}), club:{club} ',
                        color=color,
                        fill=True,
    ).add_to(kol)
    return kol