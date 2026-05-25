# -*- coding: utf-8 -*-
__title__ = "MirrorFix"
__doc__ = "Selecione as instâncias erradas e clique aqui para corrigir."
 
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
 
FAMILIAS_PERMITIDAS = ["TAG23-E-Eletr-Fiação", "HDR23-E-Eletr-Fiação"]
 
doc   = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view  = doc.ActiveView
 
class LineGridFilter(ISelectionFilter):
    def AllowElement(self, el):
        return isinstance(el, (Grid, DetailLine, ModelLine, CurveElement))
    def AllowReference(self, ref, point):
        return False
 
ids = uidoc.Selection.GetElementIds()
 
if not ids:
    TaskDialog.Show("Erro", "Selecione as famílias espelhadas.")
else:
    ids_filtrados = []
    for eid in ids:
        el = doc.GetElement(eid)
        if not isinstance(el, FamilyInstance):
            continue
        try:
            family_name = el.Symbol.Family.Name
            if any(nome in family_name for nome in FAMILIAS_PERMITIDAS):
                ids_filtrados.append(eid)
        except:
            continue
 
    if not ids_filtrados:
        TaskDialog.Show("Aviso", "Nenhuma família permitida encontrada na seleção.")
    else:
        try:
            ref = uidoc.Selection.PickObject(ObjectType.Element, LineGridFilter(), "Clique no grid ou linha de eixo de simetria")
            eixo_el = doc.GetElement(ref.ElementId)
        except:
            TaskDialog.Show("Cancelado", "Nenhum eixo selecionado.")
            import sys; sys.exit()
 
        if isinstance(eixo_el, Grid):
            eixo_curve = eixo_el.Curve
        else:
            eixo_curve = eixo_el.GeometryCurve
 
        eixo_dir = (eixo_curve.GetEndPoint(1) - eixo_curve.GetEndPoint(0)).Normalize()
        eixo_horizontal = abs(eixo_dir.X) > abs(eixo_dir.Y)
 
        t = Transaction(doc, "Corrigir espelhamento")
        t.Start()
 
        corrigidos = 0
        ignorados = len(ids) - len(ids_filtrados)
        erros = []
 
        for eid in ids_filtrados:
            el = doc.GetElement(eid)
            old_id = eid.IntegerValue
 
            try:
                loc = el.Location
                if not isinstance(loc, LocationPoint):
                    erros.append("ID {}: sem LocationPoint".format(old_id))
                    continue
 
                loc_pt = loc.Point
 
                bb = el.get_BoundingBox(view)
                if bb is None:
                    bb = el.get_BoundingBox(None)
 
                if eixo_horizontal:
                    centro_y = (bb.Min.Y + bb.Max.Y) / 2
                    off_y = loc_pt.Y - centro_y
                    ponto = XYZ(loc_pt.X, centro_y - off_y, loc_pt.Z)
                else:
                    centro_x = (bb.Min.X + bb.Max.X) / 2
                    off_x = loc_pt.X - centro_x
                    ponto = XYZ(centro_x - off_x, loc_pt.Y, loc_pt.Z)
 
                params_backup = {}
                for param in el.Parameters:
                    if param.IsReadOnly or not param.HasValue:
                        continue
                    nome = param.Definition.Name
                    if param.StorageType == StorageType.String:
                        params_backup[nome] = ("str", param.AsString())
                    elif param.StorageType == StorageType.Double:
                        params_backup[nome] = ("dbl", param.AsDouble())
                    elif param.StorageType == StorageType.Integer:
                        params_backup[nome] = ("int", param.AsInteger())
                    elif param.StorageType == StorageType.ElementId:
                        params_backup[nome] = ("eid", param.AsElementId())
 
                symbol = el.Symbol
                doc.Delete(el.Id)
 
                nova = doc.Create.NewFamilyInstance(ponto, symbol, view)
                doc.Regenerate()  # FIX: força inicialização completa antes de setar parâmetros
 
                for nome, (tipo, val) in params_backup.items():
                    pp = nova.LookupParameter(nome)
                    if pp and not pp.IsReadOnly:
                        try:
                            if tipo == "str" and val is not None:
                                pp.Set(val)
                            elif tipo == "dbl":
                                pp.Set(val)
                            elif tipo == "int":
                                pp.Set(val)
                            elif tipo == "eid":
                                pp.Set(val)
                        except Exception as e:
                            erros.append("Param '{}' ID {}: {}".format(nome, old_id, str(e)))
 
                corrigidos += 1
 
            except Exception as e:
                erros.append("ID {}: {}".format(old_id, str(e)))
 
        t.Commit()
 
        msg = "{} famílias corrigidas".format(corrigidos)
        if ignorados:
            msg += "\n{} elementos ignorados (família não permitida)".format(ignorados)
        if erros:
            msg += "\n\nErros:\n" + "\n".join(erros)
        TaskDialog.Show("Resultado", msg)