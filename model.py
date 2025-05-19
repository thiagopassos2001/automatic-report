from docxtpl import DocxTemplate

def OficioPatologia(file_path):
    template_path = "internal_data\template\Modelo_Patologia.docx"
    accidents_path = ""

    context = {
        "city_day_month_year":"",
        "count_segments":"",
        "road_name":"",
        "SRE_list":"",
        "img_ pavement_map":"",
        "img_ pavement_failure":"",
        "count_total_accidents":"",
        "count_serious_accidents":"",
        "count_fatal_accidents":""
    }

if __name__=="__main__":
    print("Hello World")