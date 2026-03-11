# -*- coding: utf-8 -*-
from pyrevit import revit, forms
from Autodesk.Revit import DB as db

doc = revit.doc
selection = revit.get_selection().elements

if not selection:
    forms.alert("Selecione os elementos no Revit primeiro!")
else:
    levels = db.FilteredElementCollector(doc).OfClass(db.Level).ToElements()
    level_dict = {l.Name: l for l in levels}
    
    dest_level_name = forms.SelectFromList.show(
        sorted(level_dict.keys()), 
        title="Mover seleção para qual nível?",
        multiselect=False
    )

    if dest_level_name:
        target_level = level_dict[dest_level_name]
        target_elevation = target_level.ProjectElevation

        t = db.Transaction(doc, "Nn: Mover Multi-Categoria")
        try:
            t.Start()
            for el in selection:
                # --- LÓGICA POR CATEGORIA ---
                category = el.Category.Name if el.Category else ""
                
                # Definir quais parâmetros buscar com base na categoria
                if "Paredes" in category or "Walls" in category:
                    p_lvl = db.BuiltInParameter.WALL_BASE_CONSTRAINT
                    p_off = db.BuiltInParameter.WALL_BASE_OFFSET
                elif "Pisos" in category or "Floors" in category:
                    p_lvl = db.BuiltInParameter.LEVEL_PARAM
                    p_off = db.BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM
                elif "Telhados" in category or "Roofs" in category:
                    p_lvl = db.BuiltInParameter.ROOF_BASE_LEVEL_PARAM
                    p_off = db.BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM
                elif "Ambientes" in category or "Rooms" in category:
                    p_lvl = db.BuiltInParameter.ROOM_LEVEL_ID
                    p_off = None # Ambientes geralmente não têm offset editável assim
                else:
                    # Genérico para móveis e outros componentes
                    p_lvl = db.BuiltInParameter.FAMILY_LEVEL_PARAM
                    p_off = db.BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM

                # --- APLICAÇÃO ---
                lvl_param = el.get_Parameter(p_lvl)
                if lvl_param and not lvl_param.IsReadOnly:
                    current_level = doc.GetElement(lvl_param.AsElementId())
                    if current_level:
                        diff = current_level.ProjectElevation - target_elevation
                        
                        # Troca o nível
                        lvl_param.Set(target_level.Id)
                        
                        # Ajusta o offset se existir para a categoria
                        if p_off:
                            off_param = el.get_Parameter(p_off)
                            if off_param and not off_param.IsReadOnly:
                                old_val = off_param.AsDouble()
                                off_param.Set(old_val + diff)
            
            t.Commit()
            forms.alert("Sucesso! Elementos de diferentes categorias foram remapeados.")
        except Exception as e:
            if t.GetStatus() == db.TransactionStatus.Started:
                t.RollBack()
            forms.alert("Erro: {}".format(str(e)))