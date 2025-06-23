import contextily as cx
import matplotlib.pyplot as plt
import matplotlib
import pandas as pd
import geopandas as gpd
import seaborn as sns
import os
import numpy as np
from docxtpl import DocxTemplate,InlineImage
from docx.shared import Mm
from datetime import datetime

CRS = "EPSG:31984"
config_set = {
        "CE-350-1":(4000,'4 km',125,300,'lower center','lower',0.3),
        "CE-531-1":(1000,'1 km',60,150,'center left','lower',0.15)
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

document_template_path = {
    "Acostamento":"bd/template/Modelo_Acostamento.docx",
    "Iluminação":"bd/template/Modelo_Iluminacao.docx",
    "Passeio":"bd/template/Modelo_Passeio.docx",
    "Patologia":"bd/template/Modelo_Patologia.docx",
    "BaiaOnibus":"bd/template/Modelo_BaiaOnibus.docx"
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
        base_shape.plot(ax=ax,linewidth=1.5,alpha=0.7,color="#593CB0",zorder=5,label='Trecho')
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

def OfficialDocument(id,gdf_path,img_path,document_type,config="auto",document_template_path=document_template_path):
    print("Iniciando processamento...")
    # Verificação do nome do template
    if document_type not in document_template_path.keys():
        raise ValueError(f"'{document_type}' não é um valor válido.")
    template = DocxTemplate(document_template_path[document_type])

    gdf = gpd.read_file(gdf_path).to_crs("EPSG:31984")

    # Verifica se o shape está vazio
    if gdf.empty:
        raise ValueError(f"O shape fornecido possui '{len(gdf)}' feições.")

    # Adiciona a coluna condição
    gdf["Condição"] = np.nan

    road_name = "CE-"+id.split("-")[1]
    root_dir = f"bd/report"
    save_file_path = f"bd/report/{id} Ofício {document_type}.docx"

    if config=="auto":
        config = config_set[id]

    date_str = datetime.today().strftime('%d-%m-%Y').split("-")
    date_str = date_str[0]+" de "+f"{month_date[date_str[1]]}"+" de "+date_str[-1]
    
    accidents_path = "bd/support/Sinistros Consolidados (2022 - 2024) - PSV.xlsx"
    accidents_name = "sinistros_22-24"
    base_map_path = "bd/support/Shape_SRE_15_04_2025_Compatibilizado.gpkg"
    src_psv_path = "bd/support/base.csv"
    
    df_sre = pd.read_csv(src_psv_path)
    df_sre = df_sre[df_sre["ID_PSV"]==id]
    sre_list = df_sre["SRE"].tolist()

    gdf_sre = gpd.read_file(base_map_path).to_crs(CRS)
    gdf_sre = gdf_sre[gdf_sre["SRE"].isin(sre_list)]

    df_accidents = pd.read_excel(accidents_path,sheet_name=accidents_name)
    df_accidents = df_accidents[df_accidents["SRE"].isin(sre_list)]

    map_img_path = os.path.join(root_dir,f"{id}_img_map_{document_type.lower()}.png")
    fig, ax = NewMap([gdf],"Condição",config=config,base_shape=gdf_sre)
    ax.legend(['Trecho','Trecho Crítico'],loc=config[4])
    plt.savefig(map_img_path, bbox_inches='tight')
    plt.clf()

    context = {
        "city_day_month_year":f"Fortaleza, {date_str}",
        "count_segments":str(len(sre_list)),
        "road_name":road_name,
        "SRE_list":", ".join(sre_list[:-1])+" e "+sre_list[-1] if len(sre_list)>1 else sre_list[0],
        "img_map":InlineImage(template,map_img_path,width=Mm(160)),
        "img":InlineImage(template,img_path,width=Mm(120)),
        "count_total_accidents":str(len(df_accidents)),
        "count_serious_accidents":str(len(df_accidents[df_accidents["gravidade"].isin(["Grave","GRAVE","Leve","LEVE"])])),
        "count_fatal_accidents":str(len(df_accidents[df_accidents["gravidade"].isin(["Fatal","FATAL"])])),
    }   

    template.render(context)
    template.save(save_file_path)

    gdf[["geometry"]].to_file(os.path.join(root_dir,f"{id.replace('-','_')}_{document_type.lower()}.kml"),driver="KML")

    print(f"Ofício salvo em {save_file_path}")

if __name__=="__main__":
    id = "CE-531-1"
    shape_path = r"\\192.168.0.5\tecnico1\TRABALHOS\ANDAMENTO\2025-CE-EST-DET-EPROS\3. PRODUTOS\2025 - 72 - Projeto Trechos Críticos (PSV)\04. PRODUTOS\CE-531-1\4 - OFÍCIOS\Shape\gpkg\Patologia.gpkg"
    img_path = r"\\192.168.0.5\tecnico1\TRABALHOS\ANDAMENTO\2025-CE-EST-DET-EPROS\3. PRODUTOS\2025 - 72 - Projeto Trechos Críticos (PSV)\04. PRODUTOS\CE-531-1\4 - OFÍCIOS\Imagens\Patologia.JPG"
    
    # (scale_len,scale_text,thickness,vert_offset,loc_legend,loc_bar_y,arrow_length)
    config = "auto"

    result = OfficialDocument(id,shape_path,img_path,"Patologia",config=config)