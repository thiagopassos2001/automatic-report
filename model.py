import contextily as cx
import matplotlib.pyplot as plt
import matplotlib
import pandas as pd
import geopandas as gpd
import seaborn as sns
import os
from docxtpl import DocxTemplate,InlineImage
from docx.shared import Mm
from datetime import datetime

CRS = "EPSG:31984"
config_set = {
        "M1-S01-CE-350-1":(4000,'4 km',125,300,'lower center','lower',0.3)
    }
month_date = {
    '01':'janeiro',
    '02':'fevereiro',
    '03':'março',
    '04':'abril',
    '05':'maio',
    '06':'junho',
    '07':'julho',
    '08':'agosto',
    '09':'setembro',
    '10':'outubro',
    '11':'novembro',
    '12':'dezembro'
    }

sns.set_theme()
plt.rcParams['font.family'] = 'sans serif'

def ConcatList(list_list):
    list_concat = []

    for i in list_list:
        list_concat = list_concat + i

    return list_concat

def NewMap(
        gdf_list,
        label_column,
        color_set=['red','green','blue','orange','cyan','yellow','purple','pink','lime','olive','gold'],
        config='auto',
        base_shape=None):
    
    if len(gdf_list)>len(color_set):
        raise ValueError('Quantidade de "color_set" insuficiente!')

    # Imagens
    fig, ax = plt.subplots(figsize=(10,10),dpi=600)

    if type(base_shape)==gpd.GeoDataFrame:
        base_shape.plot(ax=ax,linewidth=1,alpha=0.7,color='#a080ff',zorder=5,label='Trecho')
        gdf = pd.concat(gdf_list+[base_shape],ignore_index=True)
    else:
        gdf = pd.concat(gdf_list,ignore_index=True)

    gdf = gpd.GeoDataFrame(gdf,geometry='geometry',crs='EPSG:31984')

    # Plot
    for i,j in zip(gdf_list,color_set[:len(gdf_list)]):
        i.plot(ax=ax,linewidth=2,label=i[label_column].iloc[0],color=j,zorder=15,alpha=0.5)

    # Correção para rodovias mais verticais
    x_min = min([j[0] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])])
    x_max = max([j[0] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])])
    y_min = min([j[1] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])])
    y_max = max([j[1] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])])

    if ((16/9)*(y_max-y_min))>(x_max-x_min):
        w = (y_max-y_min)*(16/9)
        left_lim = ((x_max+x_min)*0.5)-(w*0.5)
        right_lim = ((x_max+x_min)*0.5)+(w*0.5)
        ax.set_xlim([left_lim, right_lim])

        # Barra escala
        x = left_lim + w*0.025
    else:
        # Barra escala
        x = min([j[0] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])])

    map_src = 'https://worldtiles3.waze.com/tiles/{z}/{x}/{y}.png'
    # map_src = 'https://services.arcgisonline.com/ArcGIS/rest/services/Canvas/World_Light_Gray_Base/MapServer/tile/{z}/{y}/{x}'
    # map_src = "https://mt1.google.com/vt/lyrs%3Dy%26x%3D{x}%26y%3D{y}%26z%3D{z}"
    cx.add_basemap(ax, crs=gdf.crs,source=map_src)
    # Remover números das bordas
    ax.get_xaxis().set_ticks([])
    ax.get_yaxis().set_ticks([])

    # Barra de Escala
    lon_size = max([j[0] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])]) - min([j[0] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])])
    if config=='auto':
        scale_len = lon_size//10
        if scale_len<1000:
            scale_text = f'{int(scale_len)} m'
        else:
            scale_text = f'{int(scale_len//1000)} km'

        thickness = scale_len//10
        vert_offset = scale_len/20
        loc_legend = 'upper left'
        loc_bar_y = 'lower'
        arrow_length = 0.15
    else:
        scale_len = config[0]
        scale_text = config[1]
        thickness = config[2]
        vert_offset = config[3]
        loc_legend = config[4]
        loc_bar_y = config[5]
        arrow_length = config[6]

    if loc_bar_y=='lower':
        y = min([j[1] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])])
    elif loc_bar_y=='upper':
        y = max([j[1] for j in ConcatList([list(i.coords) for i in gdf.explode().geometry])]) - 2*vert_offset
    else:
        pass

    scale_rect = matplotlib.patches.Rectangle((x,y),scale_len,thickness,linewidth=0.5,edgecolor='k',facecolor='k')
    ax.add_patch(scale_rect)
    plt.text(int(x+scale_len/2),int(y+vert_offset),s=scale_text, fontsize=12, horizontalalignment='center')

    # Adicionar norte
    x, y = 0.95, 0.95
    ax.annotate('N', xy=(x,y), xytext=(x,y-arrow_length),
                arrowprops=dict(facecolor='black',width=5,headwidth=15),
                ha='center',va='center',fontsize=15,
                xycoords=ax.transAxes)

    plt.legend(loc=loc_legend)

    return fig, ax

def OficioPatologia(
        file_path,
        config="auto",
        config_set=config_set,
        CRS=CRS,
        month_date=month_date):
    
    id = "-".join(os.path.basename(file_path).split("-")[:5])
    road_name = "CE-"+id.split("-")[3]
    root_dir = f"internal_data/report/{id}"
    save_file_path = os.path.join(root_dir,f"Ofício {id} - Patologia.docx")

    if config=="auto":
        config = config_set[id]

    date_str = datetime.today().strftime('%d-%m-%Y').split("-")
    date_str = date_str[0]+" de "+f"{month_date[date_str[1]]}"+" de "+date_str[-1]
    
    template_path = "internal_data/template/Modelo_Patologia.docx"
    accidents_path = "internal_data/support/Sinistros Consolidados (2022 - 2024) - PSV.xlsx"
    accidents_name = "sinistros_22-24"
    base_map_path = "internal_data/support/Shape_SRE_15_04_2025_Compatibilizado.gpkg"
    src_psv_path = "internal_data/support/1. Acompanhamento Base.xlsx"
    sre_psv_name = "Trechos"

    map_img_path = os.path.join(root_dir,"img_pavement_failure_map.png")
    img_path =  os.path.join(root_dir,"img_pavement_failure.JPG")

    df_sre = pd.read_excel(src_psv_path,sheet_name=sre_psv_name)
    df_sre = df_sre[df_sre["ID PSV"]==id]
    sre_list = df_sre["SRE"].tolist()

    gdf_sre = gpd.read_file(base_map_path).to_crs(CRS)
    gdf_sre = gdf_sre[gdf_sre["SRE"].isin(sre_list)]

    df_accidents = pd.read_excel(accidents_path,sheet_name=accidents_name)
    df_accidents = df_accidents[df_accidents["SRE"].isin(sre_list)]

    gdf = gpd.read_file(file_path).to_crs(CRS)

    fig, ax = NewMap([gdf],"Condição",config=config,base_shape=gdf_sre)
    ax.legend(['Trecho','Ponto Crítico'],loc=config[4])
    plt.savefig(map_img_path, bbox_inches='tight')

    template = DocxTemplate(template_path)

    context = {
        "city_day_month_year":f"Fortaleza, {date_str}",
        "count_segments":str(len(sre_list)),
        "road_name":road_name,
        "SRE_list":", ".join(sre_list[:-1])+" e "+sre_list[-1] if len(sre_list)>1 else sre_list[0],
        "img_pavement_failure_map":InlineImage(template,map_img_path,width=Mm(160)),
        "img_pavement_failure":InlineImage(template,img_path,width=Mm(120)),
        "count_total_accidents":str(len(df_accidents)),
        "count_serious_accidents":str(len(df_accidents[df_accidents["gravidade"].isin(["Grave","GRAVE","Leve","LEVE"])])),
        "count_fatal_accidents":str(len(df_accidents[df_accidents["gravidade"].isin(["Fatal","FATAL"])])),
    }
    
    plt.clf()

    template.render(context)
    template.save(save_file_path)

    print(f"Ofício Patologia salvo em: {save_file_path}")

def OficioIluminacao(
        file_path,
        config="auto",
        config_set=config_set,
        CRS=CRS,
        month_date=month_date):

    id = "-".join(os.path.basename(file_path).split("-")[:5])
    road_name = "CE-"+id.split("-")[3]
    root_dir = f"internal_data/report/{id}"
    save_file_path = os.path.join(root_dir,f"Ofício {id} - Iluminação.docx")

    if config=="auto":
        config = config_set[id]

    date_str = datetime.today().strftime('%d-%m-%Y').split("-")
    date_str = date_str[0]+" de "+f"{month_date[date_str[1]]}"+" de "+date_str[-1]
    
    template_path = "internal_data/template/Modelo_Iluminacao.docx"
    accidents_path = "internal_data/support/Sinistros Consolidados (2022 - 2024) - PSV.xlsx"
    accidents_name = "sinistros_22-24"
    base_map_path = "internal_data/support/Shape_SRE_15_04_2025_Compatibilizado.gpkg"
    src_psv_path = "internal_data/support/1. Acompanhamento Base.xlsx"
    sre_psv_name = "Trechos"

    map_img_path = os.path.join(root_dir,"img_public_lighting_failure_map.png")
    img_path =  os.path.join(root_dir,"img_public_lighting_failure.JPG")

    df_sre = pd.read_excel(src_psv_path,sheet_name=sre_psv_name)
    df_sre = df_sre[df_sre["ID PSV"]==id]
    sre_list = df_sre["SRE"].tolist()

    gdf_sre = gpd.read_file(base_map_path).to_crs(CRS)
    gdf_sre = gdf_sre[gdf_sre["SRE"].isin(sre_list)]

    df_accidents = pd.read_excel(accidents_path,sheet_name=accidents_name)
    df_accidents = df_accidents[df_accidents["SRE"].isin(sre_list)]

    gdf = gpd.read_file(file_path).to_crs(CRS)

    fig, ax = NewMap([gdf],"Condição",config=config,base_shape=gdf_sre)
    ax.legend(['Trecho','Ponto Crítico'],loc=config[4])
    plt.savefig(map_img_path, bbox_inches='tight')

    template = DocxTemplate(template_path)

    context = {
        "city_day_month_year":f"Fortaleza, {date_str}",
        "count_segments":str(len(sre_list)),
        "road_name":road_name,
        "SRE_list":", ".join(sre_list[:-1])+" e "+sre_list[-1] if len(sre_list)>1 else sre_list[0],
        "img_public_lighting_failure_map":InlineImage(template,map_img_path,width=Mm(160)),
        "img_public_lighting_failure":InlineImage(template,img_path,width=Mm(120)),
        "count_total_accidents":str(len(df_accidents)),
        "count_serious_accidents":str(len(df_accidents[df_accidents["gravidade"].isin(["Grave","GRAVE","Leve","LEVE"])])),
        "count_fatal_accidents":str(len(df_accidents[df_accidents["gravidade"].isin(["Fatal","FATAL"])])),
    }
    
    plt.clf()

    template.render(context)
    template.save(save_file_path)

    print(f"Ofício Iluminação salvo em: {save_file_path}")

if __name__=="__main__":
    OficioPatologia("internal_data/shape/M1-S01-CE-350-1-PAT.gpkg")
    OficioIluminacao("internal_data/shape/M1-S01-CE-350-1-ILU.gpkg")
    OficioIluminacao("internal_data/shape/M1-S01-CE-350-1-ACO.gpkg")