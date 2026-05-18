# -*- coding: utf-8 -*-
from pyrevit import revit, forms
from Autodesk.Revit import DB as db
from System.Collections.Generic import List

doc = revit.doc
view = doc.ActiveView

# 1. Pegar todos os nÃ­veis do projeto para o usuÃ¡rio escolher
levels = db.FilteredElementCollector(doc).OfClass(db.Level).ToElements()
level_dict = {l.Name: l for l in levels}

selected_level_name = forms.SelectFromList.show(
    sorted(level_dict.keys()), 
    title="Selecione o nÃ­vel para isolar elementos",
    multiselect=False
)

if selected_level_name:
    target_level_id = level_dict[selected_level_name].Id
    
    # 2. Coletar elementos na vista ativa
    all_elements = db.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType().ToElements()
    
    ids_to_select = []
    
    # Lista de todos os parÃ¢metros internos que o Revit usa para nÃ­veis
    level_params = [
        db.BuiltInParameter.FAMILY_LEVEL_PARAM,             # Componentes, MobiliÃ¡rio
        db.BuiltInParameter.LEVEL_PARAM,                    # Pisos, Telhados
        db.BuiltInParameter.WALL_BASE_CONSTRAINT,           # Paredes
        db.BuiltInParameter.ROOM_LEVEL_ID,                  # Ambientes
        db.BuiltInParameter.SCHEDULE_LEVEL_PARAM,           # VÃ¡rios elementos
        db.BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM, # Colunas, Vigas
        db.BuiltInParameter.ROOF_BASE_LEVEL_PARAM,          # Telhados especÃ­ficos
        db.BuiltInParameter.IMPORT_BASE_LEVEL               # VÃ­nculos e CADs
    ]

    for el in all_elements:
        found = False
        for p_id in level_params:
            param = el.get_Parameter(p_id)
            if param and param.AsElementId() == target_level_id:
                found = True
                break
        
        if found:
            ids_to_select.append(el.Id)

    # 3. Executar SeleÃ§Ã£o e Isolamento
    if ids_to_select:
        collection = List[db.ElementId](ids_to_select)
        
        # Seleciona os objetos
        revit.get_selection().set_to(collection)
        
        # Inicia a transaÃ§Ã£o para o isolamento temporÃ¡rio
        t = db.Transaction(doc, "Nn: Isolar Tudo no Nivel")
        try:
            t.Start()
            view.IsolateElementsTemporary(collection)
            t.Commit()
            forms.alert("{} elementos de diversas categorias isolados.".format(len(ids_to_select)))
        except Exception as e:
            if t.GetStatus() == db.TransactionStatus.Started:
                t.RollBack()
            forms.alert("Erro ao isolar: {}".format(str(e)))
    else:
        forms.alert("Nenhum elemento encontrado vinculado ao nÃ­vel {} nesta vista.".format(selected_level_name))

