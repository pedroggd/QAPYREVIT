# -*- coding: utf-8 -*-
from pyrevit import revit, DB
import Autodesk.Revit.Exceptions as Exceptions

# 1. Obter o documento padrão do pyRevit
doc = revit.doc
uidoc = revit.uidoc

def get_wall_face_references(wall, link_instance):
    """Extrai referências geométricas válidas das faces de uma parede linkada."""
    refs = []
    opt = DB.Options()
    opt.ComputeReferences = True # Obrigatório para poder cotar
    
    geom_elem = wall.get_Geometry(opt)
    if geom_elem:
        for geom_obj in geom_elem:
            if isinstance(geom_obj, DB.Solid) and geom_obj.Faces.Size > 0:
                for face in geom_obj.Faces:
                    # Pegar apenas faces verticais planas (para simplificar)
                    if isinstance(face, DB.PlanarFace):
                        # Converte a referência da face para uma referência de link
                        link_ref = face.Reference.CreateLinkReference(link_instance)
                        refs.append(link_ref)
    return refs

# Iniciar Transação
with revit.Transaction("Cota Automatica Itens Linkados"):
    active_view = doc.ActiveView
    
    # Pegar todos os links
    linked_instances = DB.FilteredElementCollector(doc).OfClass(DB.RevitLinkInstance).ToElements()
    
    for link_instance in linked_instances:
        linked_doc = link_instance.GetLinkDocument()
        
        if linked_doc:
            walls = DB.FilteredElementCollector(linked_doc).OfClass(DB.Wall).ToElements()
            reference_array = DB.ReferenceArray()
            
            for wall in walls:
                # 2. Puxar referências válidas das faces (não do elemento)
                faces_refs = get_wall_face_references(wall, link_instance)
                
                # Adiciona apenas uma face por parede para não embolar a cota
                if faces_refs:
                    reference_array.Append(faces_refs[0]) 
            
            # 3. Criar a cota
            if reference_array.Size >= 2:
                # ATENÇÃO: A linha da cota DEVE cruzar perpendicularmente as faces que você está cotando.
                # Aqui definimos uma linha paralela ao eixo Y (apenas para teste)
                p1 = DB.XYZ(0, 0, 0)
                p2 = DB.XYZ(0, 10, 0) 
                dim_line = DB.Line.CreateBound(p1, p2)
                
                try:
                    dimension = doc.Create.NewDimension(active_view, dim_line, reference_array)
                except Exceptions.ArgumentException as e:
                    print("Erro ao gerar cota (verifique se a linha cruza as paredes selecionadas): {}".format(e))