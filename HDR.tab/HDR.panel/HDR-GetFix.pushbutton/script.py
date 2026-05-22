# -*- coding: utf-8 -*-

__title__ = "PuxaPrancha+"
__author__ = "PyRevit Plugin"

import clr
import System
import re

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Window, Thickness, GridLength, GridUnitType
from System.Windows.Controls import (
    StackPanel, Label, TextBox, Button, ScrollViewer,
    Separator, Grid, ColumnDefinition, RowDefinition, ComboBox, ComboBoxItem,
    TabControl, TabItem, CheckBox
)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import FontWeights, ResizeMode, WindowStartupLocation
from System.Windows.Controls import ScrollBarVisibility

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, FamilySymbol, ViewSheet, Transaction,
    TransactionGroup, BuiltInCategory, BuiltInParameter,
    Viewport, XYZ, ViewType, StorageType, ElementId
)
import Autodesk.Revit.DB as DB

logger = script.get_logger()
doc    = revit.doc
uidoc  = revit.uidoc

FAMILIA_NOME = "GRAF23-Carimbo (NBR-16752) - PLAENGE"

PARAMS_CARIMBO = [
    ("Numero do Logo",  "Numero do Logo"),
    ("OBRA",            "Obra"),
    ("CLIENTE",         "Nome do Cliente"),
    ("LOCAL",           "Local"),
    ("PROJETISTA",      "Designed By"),
    ("DOCUMENTACAO",    "Drawn By"),
    ("COORDENADOR",     "Checked By"),
    ("APROVADOR",       "Approved By"),
    ("ARQ_CONSTRUTORA", "Arq. Construtora"),
    ("ARQ_PROJETISTA",  "Arq. Projetista"),
    ("EMISSAO_INICIAL", "Sheet Issue Date"),
]

def remove_accents(s):
    if not s: return ""
    mapping = {
        u'A':u'A', u'A':u'A', u'A':u'A', u'A':u'A',
        u'E':u'E', u'E':u'E', u'E':u'E',
        u'I':u'I', u'I':u'I', u'I':u'I',
        u'O':u'O', u'O':u'O', u'O':u'O', u'O':u'O',
        u'U':u'U', u'U':u'U', u'U':u'U',
        u'C':u'C',
        u'\xe1':u'a', u'\xe0':u'a', u'\xe3':u'a', u'\xe2':u'a',
        u'\xe9':u'e', u'\xe8':u'e', u'\xea':u'e',
        u'\xed':u'i', u'\xec':u'i', u'\xee':u'i',
        u'\xf3':u'o', u'\xf2':u'o', u'\xf5':u'o', u'\xf4':u'o',
        u'\xfa':u'u', u'\xf9':u'u', u'\xfb':u'u',
        u'\xe7':u'c', u'\xba':u'o', u'\xaa':u'a',
        u'\xc1':u'A', u'\xc0':u'A', u'\xc3':u'A', u'\xc2':u'A',
        u'\xc9':u'E', u'\xc8':u'E', u'\xca':u'E',
        u'\xcd':u'I', u'\xcc':u'I', u'\xce':u'I',
        u'\xd3':u'O', u'\xd2':u'O', u'\xd5':u'O', u'\xd4':u'O',
        u'\xda':u'U', u'\xd9':u'U', u'\xdb':u'U',
        u'\xc7':u'C',
    }
    for k, v in mapping.items():
        s = s.replace(k, v)
    return s

def get_titleblock_type():
    for symbol in (FilteredElementCollector(doc)
                   .OfCategory(BuiltInCategory.OST_TitleBlocks)
                   .OfClass(FamilySymbol)):
        if symbol.Family.Name == FAMILIA_NOME:
            return symbol

    all_types = {}
    for s in (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_TitleBlocks)
              .OfClass(FamilySymbol)):
        param = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        label = "{} : {}".format(s.Family.Name, param.AsString() if param else "?")
        all_types[label] = s

    if not all_types:
        forms.alert("Nenhum carimbo encontrado no projeto.", exitscript=True)

    chosen = forms.SelectFromList.show(
        sorted(all_types.keys()),
        title="Selecione o Carimbo",
        multiselect=False,
    )
    if not chosen:
        script.exit()
    return all_types[chosen]

def get_all_views():
    tipos_validos = {
        ViewType.FloorPlan, ViewType.EngineeringPlan, ViewType.AreaPlan,
        ViewType.Detail, ViewType.Section, ViewType.Elevation,
        ViewType.CeilingPlan, ViewType.ThreeD,
    }
    views = []
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate:
            continue
        if v.ViewType not in tipos_validos:
            continue
        views.append(v)
    return sorted(views, key=lambda v: v.Name)


class ConfigWindow(Window):
    COR_AZUL   = Color.FromRgb(30, 80, 160)
    COR_VERDE  = Color.FromRgb(22, 160, 80)
    COR_CINZA  = Color.FromRgb(120, 120, 120)
    COR_BRANCO = Color.FromRgb(255, 255, 255)
    COR_FUNDO  = Color.FromRgb(245, 245, 245)

    def __init__(self, templates_dict, all_views):
        self.templates_dict  = templates_dict
        self.all_views       = all_views
        self.resultado       = None
        self.campos_carimbo  = {}
        self.checks_vistas   = []
        self._template_names = sorted(templates_dict.keys())

        self.Title         = "Puxar Detalhe"
        self.Width         = 540
        self.Height        = 700
        self.ResizeMode    = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background            = SolidColorBrush(self.COR_FUNDO)

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)
        root.RowDefinitions.Add(self._row(GridLength.Auto))
        root.RowDefinitions.Add(self._row(GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(self._row(GridLength.Auto))

        header = StackPanel()
        self._lbl(header, "Puxar Detalhe", 16, bold=True, cor=self.COR_AZUL)
        self._lbl(header, "Filtre, selecione as vistas, aplique template e pagine.", 10, cor=self.COR_CINZA)
        header.Children.Add(self._sep(10))
        Grid.SetRow(header, 0)
        root.Children.Add(header)

        self.tabs = TabControl()
        self.tabs.FontSize = 12
        self.tabs.Margin   = Thickness(0, 0, 0, 12)
        Grid.SetRow(self.tabs, 1)
        root.Children.Add(self.tabs)

        self._build_tab_vistas()
        self._build_tab_config()
        self._build_tab_carimbo()

        btn_ok = Button()
        btn_ok.Content    = "OK - Aplicar Template e Paginar"
        btn_ok.FontSize   = 12
        btn_ok.FontWeight = FontWeights.Bold
        btn_ok.Padding    = Thickness(10, 8, 10, 8)
        btn_ok.Background = SolidColorBrush(self.COR_VERDE)
        btn_ok.Foreground = SolidColorBrush(self.COR_BRANCO)
        btn_ok.Click     += self.on_ok
        Grid.SetRow(btn_ok, 2)
        root.Children.Add(btn_ok)

        self.Content = root

    def _row(self, height):
        r = RowDefinition(); r.Height = height; return r

    def _col(self, width):
        c = ColumnDefinition(); c.Width = width; return c

    def _sep(self, bottom=8):
        s = Separator(); s.Margin = Thickness(0, 0, 0, bottom); return s

    def _lbl(self, parent, text, size, bold=False, cor=None, mg=None):
        lbl = Label()
        lbl.Content  = text
        lbl.FontSize = size
        if bold: lbl.FontWeight = FontWeights.Bold
        if cor:  lbl.Foreground = SolidColorBrush(cor)
        if mg:   lbl.Margin     = Thickness(*mg) if len(mg) > 1 else Thickness(mg[0])
        if parent is not None:
            parent.Children.Add(lbl)
        return lbl

    def _build_tab_vistas(self):
        tab = TabItem()
        tab.Header = "Vistas ({})".format(len(self.all_views))

        outer = Grid()
        outer.Margin = Thickness(8)
        outer.RowDefinitions.Add(self._row(GridLength.Auto))
        outer.RowDefinitions.Add(self._row(GridLength.Auto))
        outer.RowDefinitions.Add(self._row(GridLength(1, GridUnitType.Star)))

        # Filtro
        fpnl = StackPanel()
        fpnl.Margin = Thickness(0, 0, 0, 6)
        self._lbl(fpnl, "Filtrar por nome (ex: DETALHE S0):", 10, bold=True)
        self.txt_filtro = TextBox()
        self.txt_filtro.FontSize = 11
        self.txt_filtro.Padding  = Thickness(6, 4, 6, 4)
        self.txt_filtro.TextChanged += self.on_filtro_changed
        fpnl.Children.Add(self.txt_filtro)
        Grid.SetRow(fpnl, 0)
        outer.Children.Add(fpnl)

        # Botoes
        bpnl = StackPanel()
        bpnl.Orientation = System.Windows.Controls.Orientation.Horizontal
        bpnl.Margin = Thickness(0, 0, 0, 6)

        btn_all = Button()
        btn_all.Content = "Marcar todos"
        btn_all.FontSize = 10
        btn_all.Padding = Thickness(8, 3, 8, 3)
        btn_all.Margin  = Thickness(0, 0, 6, 0)
        btn_all.Click  += self.on_marcar_todos
        bpnl.Children.Add(btn_all)

        btn_none = Button()
        btn_none.Content = "Desmarcar todos"
        btn_none.FontSize = 10
        btn_none.Padding = Thickness(8, 3, 8, 3)
        btn_none.Click  += self.on_desmarcar_todos
        bpnl.Children.Add(btn_none)

        Grid.SetRow(bpnl, 1)
        outer.Children.Add(bpnl)

        # Lista
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        self.lista_pnl = StackPanel()
        self._popular_lista(self.all_views)
        scroll.Content = self.lista_pnl
        Grid.SetRow(scroll, 2)
        outer.Children.Add(scroll)

        tab.Content = outer
        self.tabs.Items.Add(tab)

    def _popular_lista(self, views):
        self.lista_pnl.Children.Clear()
        self.checks_vistas = []
        for v in views:
            cb = CheckBox()
            cb.Content   = v.Name
            cb.IsChecked = True
            cb.FontSize  = 11
            cb.Margin    = Thickness(2, 2, 2, 2)
            cb.Tag       = v
            self.lista_pnl.Children.Add(cb)
            self.checks_vistas.append(cb)

    def on_filtro_changed(self, sender, args):
        termo = self.txt_filtro.Text.strip().upper()
        filtradas = [v for v in self.all_views if termo in v.Name.upper()] if termo else self.all_views
        self._popular_lista(filtradas)

    def on_marcar_todos(self, sender, args):
        for cb in self.checks_vistas: cb.IsChecked = True

    def on_desmarcar_todos(self, sender, args):
        for cb in self.checks_vistas: cb.IsChecked = False

    def _build_tab_config(self):
        tab = TabItem()
        tab.Header = "Configuracoes"

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        pnl = StackPanel()
        pnl.Margin = Thickness(12)

        self._lbl(pnl, "Template de Vista a aplicar:", 11, bold=True, mg=(0,0,0,4))
        self.combo_template = ComboBox()
        self.combo_template.FontSize = 11
        self.combo_template.Margin   = Thickness(0, 0, 0, 16)
        for nome in self._template_names:
            item = ComboBoxItem()
            item.Content = nome
            self.combo_template.Items.Add(item)
        self.combo_template.SelectedIndex = 0
        pnl.Children.Add(self.combo_template)

        self._lbl(pnl, "Escala (ex: 25, 50, 100 — vazio = nao alterar):", 11, bold=True, mg=(0,0,0,4))
        self.txt_escala = TextBox()
        self.txt_escala.FontSize = 12
        self.txt_escala.Padding  = Thickness(6, 4, 6, 4)
        self.txt_escala.Text     = "25"
        pnl.Children.Add(self.txt_escala)

        scroll.Content = pnl
        tab.Content    = scroll
        self.tabs.Items.Add(tab)

    def _build_tab_carimbo(self):
        tab = TabItem()
        tab.Header = "Carimbo"

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        outer = StackPanel()
        outer.Margin = Thickness(12)

        self._lbl(outer, "Parametros do carimbo", 12, bold=True, mg=(0,0,0,4))
        self._lbl(outer, "Aplicados em todas as pranchas criadas.", 10, cor=self.COR_CINZA, mg=(0,0,0,10))
        outer.Children.Add(self._sep(8))

        for param_name, label_text in PARAMS_CARIMBO:
            g = Grid()
            g.Margin = Thickness(0, 0, 0, 6)
            g.ColumnDefinitions.Add(self._col(GridLength(150)))
            g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
            lbl = self._lbl(None, label_text, 11, mg=(0, 4, 8, 0))
            Grid.SetColumn(lbl, 0)
            g.Children.Add(lbl)
            txt = TextBox()
            txt.FontSize = 11
            txt.Padding  = Thickness(6, 4, 6, 4)
            Grid.SetColumn(txt, 1)
            g.Children.Add(txt)
            outer.Children.Add(g)
            self.campos_carimbo[label_text] = txt

        scroll.Content = outer
        tab.Content    = scroll
        self.tabs.Items.Add(tab)

    def on_ok(self, sender, args):
        vistas_selecionadas = [cb.Tag for cb in self.checks_vistas if cb.IsChecked]
        if not vistas_selecionadas:
            forms.alert("Selecione ao menos uma vista.")
            return

        t_idx = self.combo_template.SelectedIndex
        template_id = (
            self.templates_dict[self._template_names[t_idx]]
            if 0 <= t_idx < len(self._template_names)
            else ElementId.InvalidElementId
        )

        escala_txt = self.txt_escala.Text.strip()
        try:
            escala = int(escala_txt) if escala_txt else None
        except:
            escala = None

        dados_carimbo = {}
        for param_name, txt in self.campos_carimbo.items():
            valor = txt.Text.strip()
            if valor:
                dados_carimbo[param_name] = valor

        self.resultado = {
            "vistas":        vistas_selecionadas,
            "template_id":   template_id,
            "escala":        escala,
            "dados_carimbo": dados_carimbo,
        }
        self.Close()


def executar():
    tb_type = get_titleblock_type()
    if not tb_type:
        script.exit()
    tb_type_id = tb_type.Id

    all_views = get_all_views()
    if not all_views:
        forms.alert("Nenhuma vista encontrada no projeto.", exitscript=True)

    templates = {"(Nenhum)": ElementId.InvalidElementId}
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate:
            templates[v.Name] = v.Id

    dlg = ConfigWindow(templates, all_views)
    dlg.ShowDialog()

    if not dlg.resultado:
        script.exit()

    vistas_sel    = dlg.resultado["vistas"]
    template_id   = dlg.resultado["template_id"]
    escala        = dlg.resultado["escala"]
    dados_carimbo = dlg.resultado["dados_carimbo"]

    erros          = []
    vistas_geradas = []

    with TransactionGroup(doc, "Puxar Detalhe - Paginar Vistas") as tg:
        tg.Start()

        for v in vistas_sel:
            p_sheet_num = v.get_Parameter(BuiltInParameter.VIEWER_SHEET_NUMBER)
            if p_sheet_num and p_sheet_num.AsString() and p_sheet_num.AsString() != "---":
                continue

            with Transaction(doc, "Aplicar Template: {}".format(v.Name)) as t:
                t.Start()
                try:
                    if template_id != ElementId.InvalidElementId:
                        v.ViewTemplateId = template_id
                        doc.Regenerate()

                    if escala:
                        v.Scale = escala
                        doc.Regenerate()

                    vistas_geradas.append({
                        "id": v.Id,
                    })
                    t.Commit()
                except Exception as e:
                    erros.append("Erro ao processar '{}': {}".format(v.Name, str(e)))
                    t.RollBack()

        if not vistas_geradas:
            tg.RollBack()
            script.exit()

        vp_type_id = ElementId.InvalidElementId
        NOME_VP_ALVO_CLEAN = remove_accents("01. Titulo do Desenho-Com Escala (NBR-6492) 2").upper()
        for vpt in FilteredElementCollector(doc).OfClass(DB.ElementType):
            try:
                if vpt.FamilyName == "Viewport":
                    p_nome = vpt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    if p_nome:
                        nome_clean = remove_accents(p_nome.AsString() or "").upper()
                        if nome_clean == NOME_VP_ALVO_CLEAN or (
                            "NBR-6492" in nome_clean and "ESCALA" in nome_clean and nome_clean.startswith("01.")
                        ):
                            vp_type_id = vpt.Id
                            break
            except:
                continue

        with Transaction(doc, "Paginacao de Pranchas") as t:
            t.Start()
            doc.Regenerate()

            mm_to_ft    = 1.0 / 304.8
            margem_esq  = 25.0  * mm_to_ft
            margem_dir  = 175.0 * mm_to_ft
            margem_sup  = 10.0  * mm_to_ft
            margem_inf  = 15.0  * mm_to_ft
            espacamento        = 15.0 * mm_to_ft   # espaçamento entre vistas
            DISTANCIA_TITULO_FT = -50.0 * mm_to_ft  # positivo → título desce
            MARGEM_LINHA        = 20.0 * mm_to_ft  # margem vertical uniforme

            def nova_prancha():
                sh = ViewSheet.Create(doc, tb_type_id)
                doc.Regenerate()
                tbs_sh = (FilteredElementCollector(doc, sh.Id)
                          .OfCategory(BuiltInCategory.OST_TitleBlocks)
                          .WhereElementIsNotElementType().ToElements())
                tb_instance = tbs_sh[0] if tbs_sh else None

                for param_name, valor in dados_carimbo.items():
                    p = sh.LookupParameter(param_name)
                    if not p and tb_instance:
                        p = tb_instance.LookupParameter(param_name)
                    if not p:
                        p = doc.ProjectInformation.LookupParameter(param_name)
                    if p and not p.IsReadOnly:
                        if p.StorageType == StorageType.String:
                            p.Set(str(valor))
                        elif p.StorageType == StorageType.Integer:
                            try: p.Set(int(valor))
                            except: pass
                        elif p.StorageType == StorageType.Double:
                            try: p.Set(float(valor))
                            except: pass

                if not tbs_sh:
                    raise Exception("Prancha sem carimbo.")
                bb_sh = tb_instance.get_BoundingBox(sh)
                if not bb_sh:
                    raise Exception("BoundingBox None.")
                return (
                    sh,
                    bb_sh.Min.X + margem_esq,
                    bb_sh.Max.X - margem_dir,
                    bb_sh.Min.Y + margem_inf,
                    bb_sh.Max.Y - margem_sup,
                )

            sheet, ux_min, ux_max, uy_min, uy_max = nova_prancha()

            # ----------------------------------------------------------------
            # PASSO 1 — Criar TODOS os viewports em posição provisória (0,0)
            #           e ler dimensões REAIS de cada um antes de montar o layout.
            # ----------------------------------------------------------------
            vp_infos = []
            for info in vistas_geradas:
                v = doc.GetElement(info["id"])
                if not Viewport.CanAddViewToSheet(doc, sheet.Id, v.Id):
                    erros.append("Vista '{}' nao pode ser adicionada a prancha.".format(v.Name))
                    vp_infos.append(None)
                    continue

                vp = Viewport.Create(doc, sheet.Id, v.Id, XYZ(0, 0, 0))
                doc.Regenerate()

                if vp_type_id != ElementId.InvalidElementId:
                    vp.ChangeTypeId(vp_type_id)
                    doc.Regenerate()

                outline = vp.GetBoxOutline()
                real_w  = outline.MaximumPoint.X - outline.MinimumPoint.X
                real_h  = outline.MaximumPoint.Y - outline.MinimumPoint.Y

                vp_infos.append({
                    "vp":      vp,
                    "view_id": v.Id,   # guardado aqui — Viewport nao expoe .View no IronPython
                    "real_w":  real_w,
                    "real_h":  real_h,
                    "nome":    v.Name,
                })

            # ----------------------------------------------------------------
            # PASSO 2 — Agrupar em linhas usando largura REAL de cada viewport.
            # ----------------------------------------------------------------
            linhas      = []
            linha_atual = []
            tmp_x       = ux_min

            for vp_info in vp_infos:
                if vp_info is None:
                    continue
                rw = vp_info["real_w"]
                if tmp_x + rw > ux_max and tmp_x > ux_min:
                    linhas.append(linha_atual)
                    linha_atual = [vp_info]
                    tmp_x = ux_min + rw + espacamento
                else:
                    linha_atual.append(vp_info)
                    tmp_x += rw + espacamento

            if linha_atual:
                linhas.append(linha_atual)

            # ----------------------------------------------------------------
            # PASSO 3 — Posicionar cada viewport usando dimensoes reais.
            #           Margem entre linhas uniforme (MARGEM_LINHA).
            # ----------------------------------------------------------------
            cur_y = uy_max

            for linha in linhas:
                rh_max = max(vi["real_h"] for vi in linha)

                # Verifica quebra de prancha (mesma constante MARGEM_LINHA)
                if cur_y - (rh_max + MARGEM_LINHA) < uy_min:
                    sheet, ux_min, ux_max, uy_min, uy_max = nova_prancha()
                    cur_y = uy_max

                    # Vistas desta linha precisam ser recriadas na nova prancha.
                    for vi in linha:
                        view_id = vi["view_id"]  # usa ID guardado — .View quebra no IronPython
                        try:
                            doc.Delete(vi["vp"].Id)  # remove viewport provisorio da prancha anterior
                            doc.Regenerate()
                        except:
                            pass
                        new_vp = Viewport.Create(doc, sheet.Id, view_id, XYZ(0, 0, 0))
                        doc.Regenerate()
                        if vp_type_id != ElementId.InvalidElementId:
                            new_vp.ChangeTypeId(vp_type_id)
                            doc.Regenerate()
                        outline       = new_vp.GetBoxOutline()
                        vi["vp"]      = new_vp
                        vi["real_w"]  = outline.MaximumPoint.X - outline.MinimumPoint.X
                        vi["real_h"]  = outline.MaximumPoint.Y - outline.MinimumPoint.Y

                    rh_max = max(vi["real_h"] for vi in linha)

                baseline_y = cur_y - rh_max
                cur_x      = ux_min

                for vi in linha:
                    real_w = 0.0  # inicializado antes do try para evitar uso de valor anterior
                    try:
                        vp     = vi["vp"]
                        real_w = vi["real_w"]
                        real_h = vi["real_h"]

                        outline = vp.GetBoxOutline()
                        shift_x = cur_x - outline.MinimumPoint.X
                        shift_y = baseline_y - outline.MinimumPoint.Y

                        DB.ElementTransformUtils.MoveElement(doc, vp.Id, XYZ(shift_x, shift_y, 0))
                        doc.Regenerate()

                        # Titulo posicionado abaixo da vista (DISTANCIA_TITULO_FT positivo = abaixo)
                        vp.LabelOffset = XYZ(0.0, -(real_h / 2.0) - DISTANCIA_TITULO_FT, 0)
                        doc.Regenerate()

                    except Exception as e:
                        erros.append("Erro ao posicionar vista '{}': {}".format(vi["nome"], str(e)))

                    cur_x += real_w + espacamento

                # Desce para a proxima linha com margem uniforme (MARGEM_LINHA)
                cur_y -= (rh_max + MARGEM_LINHA)

            t.Commit()
        tg.Assimilate()

    msg = "Concluido! {} vistas processadas e paginadas.".format(len(vistas_geradas))
    if erros:
        msg += "\n\nAvisos ({}):\n{}".format(len(erros), "\n".join(erros))
    forms.alert(msg)


if __name__ == '__main__':
    executar()