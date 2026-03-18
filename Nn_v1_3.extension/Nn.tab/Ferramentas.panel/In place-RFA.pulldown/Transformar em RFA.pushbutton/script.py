# -*- coding: utf-8 -*-
'''
NnBim: In-Place para RFA (Pro)
DESCRICAO:
Converte elementos DirectShape ou In-Place em familias .rfa carregaveis.
O script extrai a geometria, transporta para um template (.rft) e 
reinsere no projeto na mesma posicao.

COMO USAR:
1. Selecione o elemento (In-Place ou DirectShape).
2. Execute o script.
3. Digite o nome da nova familia.
4. Selecione o template (.rft) desejado.
5. Escolha se deseja deletar o original apos a conversao.
'''

__title__ = 'Transformar\nem RFA'
__author__ = 'NnBim Dev'

import clr
import os
import re
import tempfile
from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

# Configuracao de contexto
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

def sanitize_filename(name):
    '''Remove caracteres proibidos para nomes de arquivo.'''
    return re.sub(r'[\\/*?:"<>|]', "", name)

def get_solids(element):
    '''Extrai os solidos da geometria do elemento.'''
    opt = Options()
    opt.ComputeReferences = True
    geometry = element.get_Geometry(opt)
    solids = []
    
    def extract_solid(geom_obj):
        if isinstance(geom_obj, Solid) and geom_obj.Volume > 0:
            solids.append(geom_obj)
        elif isinstance(geom_obj, GeometryInstance):
            for inst_geom in geom_obj.GetInstanceGeometry():
                extract_solid(inst_geom)

    if geometry:
        for obj in geometry: 
            extract_solid(obj)
    return solids

def get_element_level(el):
    '''Busca o nivel do elemento de forma segura.'''
    try:
        if el.LevelId != ElementId.InvalidElementId:
            return el.LevelId
    except: pass
    
    lvl_param = el.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
    if lvl_param and lvl_param.AsElementId() != ElementId.InvalidElementId:
        return lvl_param.AsElementId()
        
    active_v = doc.ActiveView
    if hasattr(active_v, 'GenLevel') and active_v.GenLevel:
        return active_v.GenLevel.Id
        
    return FilteredElementCollector(doc).OfClass(Level).FirstElement().Id

def main():
    if doc.IsFamilyDocument:
        forms.alert('Use apenas em arquivos de Projeto (.rvt)', title='Aviso')
        return

    selection = revit.get_selection()
    if not selection:
        forms.alert('Selecione um elemento antes de rodar o script.', title='NnBim')
        return

    el = selection[0]
    
    # 1. Nomeacao da Familia
    default_name = 'Familia_{}'.format(el.Id.ToString())
    raw_name = forms.ask_for_string(
        default=default_name,
        prompt='Digite o nome para a nova familia RFA:',
        title='Nomear Familia'
    )
    if not raw_name: script.exit()
    family_name = sanitize_filename(raw_name)

    # 2. Opcao de Limpeza
    should_delete = forms.alert('Deseja deletar o elemento original apos a conversao?', 
                                yes=True, no=True, title='Opcao de Limpeza')

    # 3. Escolha do Template
    template_path = forms.pick_file(file_ext='rft', title='Selecione o Template da Familia')
    if not template_path: return

    level_id = get_element_level(el)

    with forms.ProgressBar(title='Convertendo...', steps=5) as pb:
        # Passo 1: Geometria
        solidos = get_solids(el)
        pb.update_progress(1)
        if not solidos:
            forms.alert('Nenhuma geometria solida encontrada.')
            return

        bbox = el.get_BoundingBox(None)
        ponto_insercao = XYZ((bbox.Max.X + bbox.Min.X)/2, (bbox.Max.Y + bbox.Min.Y)/2, bbox.Min.Z)

        # Passo 2: Novo Documento de Familia
        fam_doc = app.NewFamilyDocument(template_path)
        pb.update_progress(2)
        
        with Transaction(fam_doc, 'Criar Geometria') as t_fam:
            t_fam.Start()
            mover = Transform.CreateTranslation(-ponto_insercao)
            for s in solidos:
                try:
                    FreeFormElement.Create(fam_doc, SolidUtils.CreateTransformed(s, mover))
                except: pass
            t_fam.Commit()
        pb.update_progress(3)

        # Passo 3: Salvamento Temporario
        temp_path = os.path.join(tempfile.gettempdir(), '{}.rfa'.format(family_name))
        save_opt = SaveAsOptions()
        save_opt.OverwriteExistingFile = True
        fam_doc.SaveAs(temp_path, save_opt)
        pb.update_progress(4)
        
        # Passo 4: Carga e Insercao no Projeto
        with revit.Transaction('NnBim: Converter para RFA'):
            loaded_fam_ref = clr.Reference[Family]()
            doc.LoadFamily(temp_path, loaded_fam_ref)
            fam_doc.Close(False)

            symbol_ids = list(loaded_fam_ref.Value.GetFamilySymbolIds())
            if not symbol_ids:
                forms.alert('Erro ao carregar a familia.')
                return
                
            symb = doc.GetElement(symbol_ids[0])
            if not symb.IsActive: symb.Activate()
            
            level = doc.GetElement(level_id)
            doc.Create.NewFamilyInstance(ponto_insercao, symb, level, Structure.StructuralType.NonStructural)
            
            if should_delete:
                doc.Delete(el.Id)
        
        pb.update_progress(5)

        try: os.remove(temp_path)
        except: pass
        
    forms.toast('Conversao concluida!')

if __name__ == '__main__':
    main()