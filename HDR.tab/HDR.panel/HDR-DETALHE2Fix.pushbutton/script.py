# -*- coding: utf-8 -*-

__title__ = "Criar\nDetalhes+"
__author__ = "PyRevit Plugin"

import clr
import System
import math
import re

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Window, Thickness, GridLength, GridUnitType
from System.Windows.Controls import (
    StackPanel, Label, TextBox, Button, ScrollViewer,
    Separator, Grid, ColumnDefinition, RowDefinition, ComboBox, ComboBoxItem,
    TabControl, TabItem
)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import FontWeights, ResizeMode, WindowStartupLocation
from System.Windows.Controls import ScrollBarVisibility

from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, FamilySymbol, ViewFamilyType, ViewFamily,
    ViewSheet, Transaction, TransactionGroup, BuiltInCategory, BuiltInParameter,
    ViewDuplicateOption, Viewport, XYZ,
    ViewType, StorageType, BoundingBoxXYZ, ElementId, Options, Line, GeometryInstance
)
import Autodesk.Revit.DB as DB
import Autodesk.Revit.Exceptions as Exceptions
from Autodesk.Revit.UI.Selection import PickBoxStyle

logger = script.get_logger()
doc   = revit.doc
uidoc = revit.uidoc

FAMILIA_NOME = "GRAF23-Carimbo (MEP-EBSERH)"

PARAMS_CARIMBO = [
    ("Número do Logo",  "Número do Logo"),
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
    try:
        if isinstance(s, str):
            s = s.decode('utf-8')
    except:
        pass
    mapping = {
        u'Á':u'A', u'À':u'A', u'Ã':u'A', u'Â':u'A',
        u'É':u'E', u'È':u'E', u'Ê':u'E',
        u'Í':u'I', u'Ì':u'I', u'Î':u'I',
        u'Ó':u'O', u'Ò':u'O', u'Õ':u'O', u'Ô':u'O',
        u'Ú':u'U', u'Ù':u'U', u'Û':u'U',
        u'Ç':u'C',
        u'á':u'a', u'à':u'a', u'ã':u'a', u'â':u'a',
        u'é':u'e', u'è':u'e', u'ê':u'e',
        u'í':u'i', u'ì':u'i', u'î':u'i',
        u'ó':u'o', u'ò':u'o', u'õ':u'o', u'ô':u'o',
        u'ú':u'u', u'ù':u'u', u'û':u'u',
        u'ç':u'c', u'º':u'o', u'ª':u'a'
    }
    for k, v in mapping.items():
        s = s.replace(k, v)
    return s

def sanitize_name(name):
    return re.sub(r'[\\:\{\}\[\]|;<>?\'~]', '-', name).strip()

def get_titleblock_symbols_by_family():
    """Retorna dict: familia_nome -> {tipo_label -> FamilySymbol}."""
    by_family = {}
    for s in (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_TitleBlocks)
              .OfClass(FamilySymbol)):
        param = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        tipo  = param.AsString() if param else "?"
        fam   = s.Family.Name
        if fam not in by_family:
            by_family[fam] = {}
        by_family[fam][tipo] = s
    return by_family

def get_titleblock_type():
    """Mantido por compatibilidade — retorna primeiro simbolo da familia preferida."""
    by_family = get_titleblock_symbols_by_family()
    if not by_family:
        forms.alert("Nenhum carimbo encontrado no projeto.", exitscript=True)
    if FAMILIA_NOME in by_family:
        tipos = by_family[FAMILIA_NOME]
        return list(tipos.values())[0]
    return list(list(by_family.values())[0].values())[0]

def get_callout_types():
    tipos = {}
    for ct in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if ct.ViewFamily in [ViewFamily.FloorPlan, ViewFamily.Detail]:
            param = ct.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if param:
                nome = param.AsString()
                familia = "Planta de Piso" if ct.ViewFamily == ViewFamily.FloorPlan else "Vista de Detalhe"
                label = "{} [{}]".format(nome, familia)
                tipos[label] = ct.Id
    return tipos

def get_viewport_types():
    """Retorna dict ordenado: label -> ElementId, de todos os tipos de Viewport no projeto."""
    tipos = {}
    for vpt in FilteredElementCollector(doc).OfClass(DB.ElementType):
        try:
            if vpt.FamilyName == "Viewport":
                p_nome = vpt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if p_nome:
                    nome = p_nome.AsString() or "?"
                    tipos[nome] = vpt.Id
        except:
            continue
    return tipos

def get_unique_view_name(base_name):
    existing_names = {v.Name for v in FilteredElementCollector(doc).OfClass(DB.View)}
    if base_name not in existing_names:
        return base_name
    contador = 1
    while True:
        new_name = "{} ({})".format(base_name, contador)
        if new_name not in existing_names:
            return new_name
        contador += 1

class CalloutConfigWindow(Window):
    COR_AZUL   = Color.FromRgb(30, 80, 160)
    COR_VERDE  = Color.FromRgb(22, 160, 80)
    COR_CINZA  = Color.FromRgb(120, 120, 120)
    COR_BRANCO = Color.FromRgb(255, 255, 255)
    COR_FUNDO  = Color.FromRgb(245, 245, 245)

    def __init__(self, templates_dict, callout_types_dict, titleblock_symbols, viewport_types_dict):
        self.templates_dict     = templates_dict
        self.callout_types_dict = callout_types_dict
        self.resultado          = None

        self._template_names     = sorted(templates_dict.keys())
        self._callout_type_names = sorted(callout_types_dict.keys())
        self.campos_carimbo      = {}

        self.viewport_types_dict  = viewport_types_dict
        # Nome preferido de viewport (mesma logica do script original)
        NOME_VP_PREF = remove_accents("01. Titulo do Desenho-Com Escala (NBR-6492) 2").upper()
        self._vp_type_labels = sorted(viewport_types_dict.keys())
        self._vp_type_best   = 0
        for i, lbl in enumerate(self._vp_type_labels):
            lbl_clean = remove_accents(lbl).upper()
            if lbl_clean == NOME_VP_PREF or (
                "NBR-6492" in lbl_clean and "ESCALA" in lbl_clean and lbl_clean.startswith("01.")
            ):
                self._vp_type_best = i
                break

        # Montar lista plana ordenada de tipos de folha (familia preferida primeiro)
        self._tb_labels  = []
        self._tb_symbols = {}
        def _add_family(fam_name):
            if fam_name not in titleblock_symbols:
                return
            for tipo in sorted(titleblock_symbols[fam_name].keys()):
                lbl = "{} : {}".format(fam_name, tipo)
                self._tb_labels.append(lbl)
                self._tb_symbols[lbl] = titleblock_symbols[fam_name][tipo]
        _add_family(FAMILIA_NOME)
        for fam in sorted(titleblock_symbols.keys()):
            if fam != FAMILIA_NOME:
                _add_family(fam)

        self.Title         = "Configuração de Callouts"
        self.Width         = 480
        self.MinHeight     = 650
        self.SizeToContent = System.Windows.SizeToContent.Height
        self.ResizeMode    = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background            = SolidColorBrush(self.COR_FUNDO)

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)

        root.RowDefinitions.Add(self._row(GridLength.Auto))
        root.RowDefinitions.Add(self._row(GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(self._row(GridLength.Auto))

        header_panel = StackPanel()
        self._lbl(header_panel, "Criar Detalhes+", 16, bold=True, cor=self.COR_AZUL, mg=(0,0,0,2))
        header_panel.Children.Add(self._sep(10))
        Grid.SetRow(header_panel, 0)
        root.Children.Add(header_panel)

        self.tabs = TabControl()
        self.tabs.FontSize = 12
        self.tabs.Margin = Thickness(0, 0, 0, 12)
        Grid.SetRow(self.tabs, 1)
        root.Children.Add(self.tabs)

        self._build_tab_config()
        self._build_tab_carimbo()

        btn_ok = Button()
        btn_ok.Content    = "OK — Iniciar Desenho de Callouts"
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
        r = RowDefinition()
        r.Height = height
        return r

    def _col(self, width):
        c = ColumnDefinition()
        c.Width = width
        return c

    def _sep(self, bottom=8):
        s = Separator()
        s.Margin = Thickness(0, 0, 0, bottom)
        return s

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

    def _build_tab_config(self):
        tab = TabItem()
        tab.Header = "Configurações"
        
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        
        pnl = StackPanel()
        pnl.Margin = Thickness(12)

        self._lbl(pnl, "Prefixo (ex: 'DETALHE', 'DET.'):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_prefixo = TextBox()
        self.txt_prefixo.FontSize = 12
        self.txt_prefixo.Padding  = Thickness(6, 4, 6, 4)
        self.txt_prefixo.Margin   = Thickness(0, 0, 0, 12)
        self.txt_prefixo.Text     = "DETALHE"
        pnl.Children.Add(self.txt_prefixo)

        self._lbl(pnl, "Identificador do Contador (ex: 'S', 'H' ou deixe vazio):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_identificador = TextBox()
        self.txt_identificador.FontSize = 12
        self.txt_identificador.Padding  = Thickness(6, 4, 6, 4)
        self.txt_identificador.Margin   = Thickness(0, 0, 0, 12)
        self.txt_identificador.Text     = "S"
        pnl.Children.Add(self.txt_identificador)

        g_cont = Grid()
        g_cont.Margin = Thickness(0, 0, 0, 12)
        g_cont.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g_cont.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))

        pnl_num = StackPanel()
        pnl_num.Margin = Thickness(0, 0, 4, 0)
        self._lbl(pnl_num, "Número Inicial:", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_num_inicial = TextBox()
        self.txt_num_inicial.FontSize = 12
        self.txt_num_inicial.Padding  = Thickness(6, 4, 6, 4)
        self.txt_num_inicial.Text     = "1"
        pnl_num.Children.Add(self.txt_num_inicial)
        Grid.SetColumn(pnl_num, 0)
        g_cont.Children.Add(pnl_num)

        pnl_zero = StackPanel()
        pnl_zero.Margin = Thickness(4, 0, 0, 0)
        self._lbl(pnl_zero, "Formato do Número:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_zeros = ComboBox()
        self.combo_zeros.FontSize = 11
        i1 = ComboBoxItem(); i1.Content = "1 dígito (ex: 1, 2, 3)"
        i2 = ComboBoxItem(); i2.Content = "2 dígitos (ex: 01, 02)"
        self.combo_zeros.Items.Add(i1)
        self.combo_zeros.Items.Add(i2)
        self.combo_zeros.SelectedIndex = 1
        pnl_zero.Children.Add(self.combo_zeros)
        Grid.SetColumn(pnl_zero, 1)
        g_cont.Children.Add(pnl_zero)

        pnl.Children.Add(g_cont)

        self._lbl(pnl, "Sufixo (ex: '- PAVIMENTO TÉRREO' ou vazio):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_sufixo = TextBox()
        self.txt_sufixo.FontSize = 12
        self.txt_sufixo.Padding  = Thickness(6, 4, 6, 4)
        self.txt_sufixo.Margin   = Thickness(0, 0, 0, 16)
        self.txt_sufixo.Text     = "- PAVIMENTO TÉRREO"
        pnl.Children.Add(self.txt_sufixo)
        
        pnl.Children.Add(self._sep(10))

        self._lbl(pnl, "Template de Vista para os Callouts:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_template = ComboBox()
        self.combo_template.FontSize = 11
        self.combo_template.Margin   = Thickness(0, 0, 0, 12)
        for nome in self._template_names:
            item = ComboBoxItem()
            item.Content = nome
            self.combo_template.Items.Add(item)
        self.combo_template.SelectedIndex = 0
        pnl.Children.Add(self.combo_template)

        self._lbl(pnl, "Tipo de Callout:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_callout = ComboBox()
        self.combo_callout.FontSize = 11
        self.combo_callout.Margin   = Thickness(0, 0, 0, 16)
        
        is_floor_plan = doc.ActiveView.ViewType in [ViewType.FloorPlan, ViewType.EngineeringPlan, ViewType.AreaPlan]
        best_idx = 0
        
        for i, nome in enumerate(self._callout_type_names):
            item = ComboBoxItem()
            item.Content = nome
            self.combo_callout.Items.Add(item)
            if is_floor_plan and "[Planta de Piso]" in nome:
                best_idx = i
            elif not is_floor_plan and "[Vista de Detalhe]" in nome and best_idx == 0:
                best_idx = i
                
        self.combo_callout.SelectedIndex = best_idx
        pnl.Children.Add(self.combo_callout)

        pnl.Children.Add(self._sep(10))

        self._lbl(pnl, "Tipo de Titulo da Vista (Viewport):", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_viewport = ComboBox()
        self.combo_viewport.FontSize = 11
        self.combo_viewport.Margin   = Thickness(0, 0, 0, 16)
        for lbl in self._vp_type_labels:
            item = ComboBoxItem()
            item.Content = lbl
            self.combo_viewport.Items.Add(item)
        self.combo_viewport.SelectedIndex = self._vp_type_best
        pnl.Children.Add(self.combo_viewport)

        pnl.Children.Add(self._sep(10))

        self._lbl(pnl, "Tipo de Folha (Carimbo):", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_folha = ComboBox()
        self.combo_folha.FontSize = 11
        self.combo_folha.Margin   = Thickness(0, 0, 0, 4)
        best_folha_idx = 0
        for i, lbl in enumerate(self._tb_labels):
            item = ComboBoxItem()
            item.Content = lbl
            self.combo_folha.Items.Add(item)
            if FAMILIA_NOME in lbl and best_folha_idx == 0:
                best_folha_idx = i
        self.combo_folha.SelectedIndex = best_folha_idx
        pnl.Children.Add(self.combo_folha)
        self._lbl(pnl, "A margem do carimbo sera lida automaticamente da prancha criada.",
                  9, cor=self.COR_CINZA, mg=(0, 0, 0, 16))

        scroll.Content = pnl
        tab.Content = scroll
        self.tabs.Items.Add(tab)

    def _build_tab_carimbo(self):
        tab = TabItem()
        tab.Header = "Carimbo"
        
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        
        outer = StackPanel()
        outer.Margin = Thickness(12)

        self._lbl(outer, "Parametros do carimbo", 12, bold=True, mg=(0,0,0,4))
        self._lbl(outer, "Estes valores serao aplicados em TODAS as pranchas criadas.", 10, cor=self.COR_CINZA, mg=(0,0,0,10))
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
            txt.Padding = Thickness(6, 4, 6, 4)
            Grid.SetColumn(txt, 1)
            g.Children.Add(txt)

            outer.Children.Add(g)
            self.campos_carimbo[label_text] = txt

        scroll.Content = outer
        tab.Content = scroll
        self.tabs.Items.Add(tab)

    def on_ok(self, sender, args):
        prefixo = self.txt_prefixo.Text.strip()
        identificador = self.txt_identificador.Text.strip()
        sufixo = self.txt_sufixo.Text.strip()
        
        try:
            num_inicial = int(self.txt_num_inicial.Text.strip())
        except:
            num_inicial = 1
            
        zeros = self.combo_zeros.SelectedIndex + 1

        t_idx = self.combo_template.SelectedIndex
        template_id = (
            self.templates_dict[self._template_names[t_idx]]
            if 0 <= t_idx < len(self._template_names)
            else ElementId.InvalidElementId
        )

        c_idx = self.combo_callout.SelectedIndex
        callout_type_id = (
            self.callout_types_dict[self._callout_type_names[c_idx]]
            if 0 <= c_idx < len(self._callout_type_names)
            else None
        )

        # Tipo de titulo de viewport selecionado
        vp_idx = self.combo_viewport.SelectedIndex
        if 0 <= vp_idx < len(self._vp_type_labels):
            vp_lbl         = self._vp_type_labels[vp_idx]
            viewport_type_id = self.viewport_types_dict[vp_lbl]
        else:
            viewport_type_id = ElementId.InvalidElementId

        # Tipo de folha selecionado
        f_idx = self.combo_folha.SelectedIndex
        if 0 <= f_idx < len(self._tb_labels):
            folha_lbl    = self._tb_labels[f_idx]
            folha_symbol = self._tb_symbols[folha_lbl]
        else:
            folha_symbol = None

        dados_carimbo = {}
        for param_name, txt in self.campos_carimbo.items():
            valor = txt.Text.strip()
            if valor:
                dados_carimbo[param_name] = valor

        self.resultado = {
            "prefixo":          prefixo,
            "identificador":    identificador,
            "num_inicial":      num_inicial,
            "zeros":            zeros,
            "folha_symbol":     folha_symbol,
            "viewport_type_id": viewport_type_id,
            "sufixo":           sufixo,
            "template_id":      template_id,
            "callout_type_id":  callout_type_id,
            "dados_carimbo":    dados_carimbo
        }
        self.Close()

def executar_fluxo_callout():
    active_view = doc.ActiveView
    _tipos_invalidos = {
        ViewType.DrawingSheet, ViewType.ProjectBrowser,
        ViewType.SystemBrowser, ViewType.Undefined,
        ViewType.Schedule, ViewType.Legend,
    }
    if active_view.ViewType in _tipos_invalidos:
        forms.alert(
            "Vista atual nao suporta Callouts.\nTipo detectado: {}\n\nAbra uma planta, corte, elevacao ou detalhe.".format(
                active_view.ViewType.ToString()
            ),
            exitscript=True
        )

    titleblock_symbols = get_titleblock_symbols_by_family()
    if not titleblock_symbols:
        forms.alert("Nenhum carimbo encontrado no projeto.", exitscript=True)

    templates = {"(Nenhum)": ElementId.InvalidElementId}
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate:
            templates[v.Name] = v.Id

    callout_types = get_callout_types()
    if not callout_types:
        forms.alert("Nenhum tipo de Callout (FloorPlan ou Detail) encontrado no projeto.", exitscript=True)

    viewport_types = get_viewport_types()

    dlg = CalloutConfigWindow(templates, callout_types, titleblock_symbols, viewport_types)
    dlg.ShowDialog()

    if not dlg.resultado:
        script.exit()

    prefixo         = dlg.resultado["prefixo"]
    identificador   = dlg.resultado["identificador"]
    num_inicial     = dlg.resultado["num_inicial"]
    zeros           = dlg.resultado["zeros"]
    sufixo          = dlg.resultado["sufixo"]
    
    template_id     = dlg.resultado["template_id"]
    callout_type_id = dlg.resultado["callout_type_id"]
    dados_carimbo   = dlg.resultado["dados_carimbo"]
    folha_symbol     = dlg.resultado.get("folha_symbol", None)
    viewport_type_id = dlg.resultado.get("viewport_type_id", ElementId.InvalidElementId)

    if folha_symbol is None:
        forms.alert("Nenhum tipo de folha selecionado.", exitscript=True)

    tb_type_id = folha_symbol.Id

    vistas_geradas = []
    erros_callout  = []
    contador       = 1

    caixas_desenhadas = []
    while True:
        formato_num = "{:0" + str(zeros) + "d}"
        numero_str_prompt = formato_num.format(num_inicial + contador - 1)
        contador_full = "{}{}".format(identificador, numero_str_prompt)
        
        partes_preview = [prefixo, contador_full, sufixo]
        nome_preview = " ".join([p for p in partes_preview if p])
        
        try:
            box = uidoc.Selection.PickBox(
                PickBoxStyle.Directional,
                "[{}] Desenhe o retângulo (ESC para opções)".format(nome_preview)
            )
            caixas_desenhadas.append(box)
            contador += 1
        except Exceptions.OperationCanceledException:
            if len(caixas_desenhadas) > 0:
                opcao = forms.alert(
                    "Seleção pausada. Você possui {} detalhe(s) na fila.\nO que deseja fazer?".format(len(caixas_desenhadas)),
                    options=["Finalizar e Criar Pranchas", "Desfazer o ÚLTIMO e Continuar", "Cancelar Script"]
                )
                if opcao == "Desfazer o ÚLTIMO e Continuar":
                    caixas_desenhadas.pop()
                    contador -= 1
                    continue
                elif opcao == "Finalizar e Criar Pranchas":
                    break
                else:
                    script.exit()
            else:
                script.exit()
        except Exception as e:
            erros_callout.append("Erro ao desenhar caixa {}: {}".format(contador, str(e)))
            break

    if not caixas_desenhadas:
        forms.alert("Nenhum detalhe desenhado. Operação cancelada.")
        script.exit()

    vistas_geradas = []

    with TransactionGroup(doc, "Gerador de Pranchas via Detalhe") as tg:
        tg.Start()

        for idx, box in enumerate(caixas_desenhadas):
            numero_str = formato_num.format(num_inicial + idx)
            contador_full = "{}{}".format(identificador, numero_str)
            
            partes_detalhe = [prefixo, contador_full, sufixo]
            nome_detalhe = " ".join([p for p in partes_detalhe if p])
            nome_detalhe = sanitize_name(nome_detalhe)

            novo_id    = None
            escala_final = 25
            real_crop_w = 0.0
            real_crop_h = 0.0
            tid_salvo   = ElementId.InvalidElementId

            with Transaction(doc, "Criar Detalhe: {}".format(nome_detalhe)) as t:
                t.Start()
                try:
                    min_x  = min(box.Min.X, box.Max.X)
                    max_x  = max(box.Min.X, box.Max.X)
                    min_y  = min(box.Min.Y, box.Max.Y)
                    max_y  = max(box.Min.Y, box.Max.Y)
                    crop_w = max_x - min_x
                    crop_h = max_y - min_y

                    if crop_w < 0.01 or crop_h < 0.01:
                        raise Exception("Retangulo muito pequeno.")

                    pt1 = XYZ(min_x, min_y, box.Min.Z)
                    pt2 = XYZ(max_x, max_y, box.Max.Z)

                    nova = DB.ViewSection.CreateCallout(doc, active_view.Id, callout_type_id, pt1, pt2)
                    novo_id = nova.Id
                    nova.Name = get_unique_view_name(nome_detalhe)

                    if template_id != ElementId.InvalidElementId:
                        try:
                            nova.ViewTemplateId = template_id
                            doc.Regenerate()
                        except:
                            pass

                    tid_salvo = nova.ViewTemplateId
                    if tid_salvo != ElementId.InvalidElementId:
                        nova.ViewTemplateId = ElementId.InvalidElementId
                        doc.Regenerate()

                    nova.Scale = 25

                    p_scope = nova.get_Parameter(BuiltInParameter.VIEWER_VOLUME_OF_INTEREST_CROP)
                    if p_scope and not p_scope.IsReadOnly:
                        p_scope.Set(ElementId.InvalidElementId)

                    doc.Regenerate()

                    nova.CropBoxActive  = True
                    nova.CropBoxVisible = True

                    bb    = nova.CropBox
                    inv_t = bb.Transform.Inverse

                    p1_local = inv_t.OfPoint(XYZ(min_x, min_y, 0))
                    p2_local = inv_t.OfPoint(XYZ(max_x, max_y, 0))

                    bb.Min = XYZ(min(p1_local.X, p2_local.X), min(p1_local.Y, p2_local.Y), bb.Min.Z)
                    bb.Max = XYZ(max(p1_local.X, p2_local.X), max(p1_local.Y, p2_local.Y), bb.Max.Z)
                    nova.CropBox = bb

                    p_ann = nova.get_Parameter(BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
                    if p_ann and not p_ann.IsReadOnly:
                        p_ann.Set(1)

                    doc.Regenerate()

                    if tid_salvo != ElementId.InvalidElementId:
                        try:
                            tmpl_elem = doc.GetElement(tid_salvo)
                            template_controla_crop = False
                            if tmpl_elem:
                                p_ctrl = tmpl_elem.get_Parameter(BuiltInParameter.VIEWER_CROP_REGION)
                                if p_ctrl and not p_ctrl.IsReadOnly:
                                    template_controla_crop = True
                            if not template_controla_crop:
                                nova.ViewTemplateId = tid_salvo
                                doc.Regenerate()
                        except:
                            pass

                    real_crop_w = crop_w
                    real_crop_h = crop_h
                    try:
                        if nova.CropBoxActive:
                            bb_chk = nova.CropBox
                            if bb_chk:
                                w_chk = abs(bb_chk.Max.X - bb_chk.Min.X)
                                h_chk = abs(bb_chk.Max.Y - bb_chk.Min.Y)
                                if w_chk > 0.001 and h_chk > 0.001:
                                    real_crop_w = w_chk
                                    real_crop_h = h_chk
                    except:
                        pass

                    t.Commit()
                except Exception as e:
                    erros_callout.append("Erro ao criar '{}': {}".format(nome_detalhe, str(e)))
                    t.RollBack()
                    novo_id = None

            if novo_id is None:
                continue

            with Transaction(doc, "Corrigir Annotation Crop: {}".format(nome_detalhe)) as t2:
                t2.Start()
                try:
                    nova_ref = doc.GetElement(novo_id)
                    if nova_ref is not None:
                        tid_atual = nova_ref.ViewTemplateId
                        if tid_atual != ElementId.InvalidElementId:
                            nova_ref.ViewTemplateId = ElementId.InvalidElementId
                            doc.Regenerate()

                        OFFSET_PARAM_NAMES = [
                            "Annotation Crop Offset Left",
                            "Annotation Crop Offset Right",
                            "Annotation Crop Offset Top",
                            "Annotation Crop Offset Bottom",
                        ]
                        for nome_p in OFFSET_PARAM_NAMES:
                            try:
                                p = nova_ref.LookupParameter(nome_p)
                                if p and not p.IsReadOnly:
                                    p.Set(0.0)
                            except:
                                pass

                        try:
                            sm = nova_ref.GetCropRegionShapeManager()
                            if sm is not None:
                                sm.BottomAnnotationCropOffset = 0.0
                                sm.TopAnnotationCropOffset    = 0.0
                                sm.LeftAnnotationCropOffset   = 0.0
                                sm.RightAnnotationCropOffset  = 0.0
                        except:
                            pass

                        doc.Regenerate()

                        if tid_atual != ElementId.InvalidElementId:
                            nova_ref.ViewTemplateId = tid_atual
                            doc.Regenerate()

                    t2.Commit()
                except Exception as e:
                    erros_callout.append("Annotation crop offset erro ({}): {}".format(nome_detalhe, str(e)))
                    t2.RollBack()

            vistas_geradas.append({
                "id":     novo_id,
                "crop_w": real_crop_w,
                "crop_h": real_crop_h,
                "escala": escala_final,
            })

        if not vistas_geradas:
            tg.RollBack()
            forms.alert("Nenhuma vista criada.\n" + "\n".join(erros_callout))
            script.exit()

        vp_type_id = viewport_type_id  # Definido pelo usuario no dialogo

        with Transaction(doc, "Paginacao de Pranchas") as t:
            t.Start()
            doc.Regenerate()

            mm              = 1.0 / 304.8
            margem_esq      = 25.0  * mm
            margem_sup      = 15.0  * mm
            margem_inf      = 15.0  * mm
            espacamento     = 10.0  * mm   # Espacamento horizontal entre viewports
            MARGEM_ENTRE_LINHAS = 20.0 * mm  # Espaco vertical extra ENTRE linhas (alem do titulo)
            ALTURA_TITULO   = 12.0  * mm   # Altura estimada do bloco de titulo do viewport

            def _calcular_margem_dir(sh, tb_instance, bb_sh):
                """
                Calcula margem_dir medindo a distancia entre o viewport de conteudo
                mais a direita e a borda direita da folha.
                A logica: coleta todos os viewports da prancha, ignora os que estao
                claramente dentro da area do carimbo (bonecos/legenda), e usa o
                limite direito do conteudo para inferir onde o carimbo comeca.
                Fallback: 180mm (valor seguro para qualquer formato EBSERH).
                """
                FALLBACK_MM = 180.0
                try:
                    vps_sh = (FilteredElementCollector(doc, sh.Id)
                              .OfClass(Viewport)
                              .ToElements())
                    if not vps_sh:
                        return FALLBACK_MM * mm

                    folha_max_x = bb_sh.Max.X   # borda direita da folha em pes

                    # Pega o outline de cada viewport e guarda o MaxX
                    max_x_list = []
                    for vp in vps_sh:
                        try:
                            ol = vp.GetBoxOutline()
                            max_x_list.append(ol.MaximumPoint.X)
                        except:
                            pass

                    if not max_x_list:
                        return FALLBACK_MM * mm

                    # O viewport mais a direita provavelmente e do carimbo.
                    # O segundo mais a direita e o limite da area util.
                    max_x_list.sort(reverse=True)
                    if len(max_x_list) >= 2:
                        conteudo_max_x = max_x_list[1]
                    else:
                        conteudo_max_x = max_x_list[0]

                    margem_calculada = folha_max_x - conteudo_max_x
                    # Sanitizar: valor minimo de 100mm, maximo de 500mm
                    margem_mm = margem_calculada / mm
                    if margem_mm < 100.0 or margem_mm > 500.0:
                        return FALLBACK_MM * mm
                    # Adicionar pequena folga (5mm) para nao cortar labels
                    return (margem_mm + 5.0) * mm
                except:
                    return FALLBACK_MM * mm

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

                # Regenera para garantir viewports do carimbo visiveis
                doc.Regenerate()
                margem_dir = _calcular_margem_dir(sh, tb_instance, bb_sh)

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
            #           e ler dimensões REAIS via GetBoxOutline antes de montar o layout.
            # ----------------------------------------------------------------
            # Posicao de staging: fora da folha, viewports separados para nao se sobrepor
            # O Revit usa pés internamente; 10000 ft esta bem fora de qualquer folha
            STAGING_BASE_X = -10000.0
            STAGING_STEP   = 5.0   # 5 ft de separacao entre cada viewport no staging

            vp_infos = []
            staging_x = STAGING_BASE_X

            for info in vistas_geradas:
                v = doc.GetElement(info["id"])
                if v is None:
                    vp_infos.append(None)
                    continue

                if not Viewport.CanAddViewToSheet(doc, sheet.Id, v.Id):
                    erros_callout.append("Vista '{}' nao pode ser adicionada a prancha.".format(v.Name))
                    vp_infos.append(None)
                    continue

                # Cria em posicao de staging unica para nao sobrepor com outros
                staging_pos = XYZ(staging_x, -10000.0, 0)
                vp = Viewport.Create(doc, sheet.Id, v.Id, staging_pos)
                doc.Regenerate()

                if vp_type_id != ElementId.InvalidElementId:
                    vp.ChangeTypeId(vp_type_id)
                    doc.Regenerate()

                try:
                    vp.LabelOffset = XYZ(0.0, 0.0, 0.0)
                    doc.Regenerate()
                except:
                    pass

                outline = vp.GetBoxOutline()
                real_w  = outline.MaximumPoint.X - outline.MinimumPoint.X
                real_h  = outline.MaximumPoint.Y - outline.MinimumPoint.Y

                # Guarda o offset do canto inferior esquerdo do outline em relacao
                # ao centro do viewport (posicao de criacao = staging_pos)
                # Isso permite calcular o shift correto na hora de posicionar.
                off_min_x = outline.MinimumPoint.X - staging_pos.X
                off_min_y = outline.MinimumPoint.Y - staging_pos.Y

                vp_infos.append({
                    "vp":        vp,
                    "view_id":   v.Id,
                    "real_w":    real_w,
                    "real_h":    real_h,
                    "off_min_x": off_min_x,
                    "off_min_y": off_min_y,
                    "nome":      v.Name,
                    "staging":   staging_pos,
                })

                staging_x += STAGING_STEP

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
            # PASSO 3 — Posicionar cada viewport usando dimensões reais.
            #           Margem entre linhas uniforme (MARGEM_LINHA).
            # ----------------------------------------------------------------
            cur_y = uy_max

            for linha in linhas:
                rh_max = max(vi["real_h"] for vi in linha)

                if cur_y - (rh_max + MARGEM_ENTRE_LINHAS) < uy_min:
                    sheet, ux_min, ux_max, uy_min, uy_max = nova_prancha()
                    cur_y = uy_max

                    staging_x_new = STAGING_BASE_X
                    for vi in linha:
                        view_id = vi["view_id"]
                        try:
                            doc.Delete(vi["vp"].Id)
                            doc.Regenerate()
                        except:
                            pass
                        s_pos = XYZ(staging_x_new, -10000.0, 0)
                        new_vp = Viewport.Create(doc, sheet.Id, view_id, s_pos)
                        doc.Regenerate()
                        if vp_type_id != ElementId.InvalidElementId:
                            new_vp.ChangeTypeId(vp_type_id)
                            doc.Regenerate()
                        try:
                            new_vp.LabelOffset = XYZ(0.0, 0.0, 0.0)
                            doc.Regenerate()
                        except:
                            pass
                        outline = new_vp.GetBoxOutline()
                        rw = outline.MaximumPoint.X - outline.MinimumPoint.X
                        rh = outline.MaximumPoint.Y - outline.MinimumPoint.Y
                        vi["vp"]        = new_vp
                        vi["real_w"]    = rw
                        vi["real_h"]    = rh
                        vi["off_min_x"] = outline.MinimumPoint.X - s_pos.X
                        vi["off_min_y"] = outline.MinimumPoint.Y - s_pos.Y
                        vi["staging"]   = s_pos
                        staging_x_new += STAGING_STEP

                    rh_max = max(vi["real_h"] for vi in linha)

                baseline_y = cur_y - rh_max
                cur_x      = ux_min

                for vi in linha:
                    real_w = 0.0
                    try:
                        vp      = vi["vp"]
                        real_w  = vi["real_w"]
                        real_h  = vi["real_h"]
                        off_x   = vi["off_min_x"]
                        off_y   = vi["off_min_y"]

                        # Centro atual do viewport = posicao de staging guardada
                        cur_center = vi["staging"]

                        # Queremos que outline.Min fique em (cur_x, baseline_y).
                        # outline.Min = center + (off_x, off_y)
                        # Portanto: target_center = (cur_x - off_x, baseline_y - off_y)
                        # Delta = target_center - cur_center
                        target_cx = cur_x - off_x
                        target_cy = baseline_y - off_y
                        delta = XYZ(target_cx - cur_center.X, target_cy - cur_center.Y, 0)

                        DB.ElementTransformUtils.MoveElement(doc, vp.Id, delta)
                        doc.Regenerate()

                        # Atualiza staging para posicao final (caso precise referenciar depois)
                        vi["staging"] = XYZ(target_cx, target_cy, 0)

                        vp.LabelOffset = XYZ(0.0, 0.0, 0.0)
                        doc.Regenerate()

                    except Exception as e:
                        erros_callout.append("Erro ao posicionar vista '{}': {}".format(vi["nome"], str(e)))

                    cur_x += real_w + espacamento

                cur_y -= (rh_max + MARGEM_ENTRE_LINHAS)

            t.Commit()
        tg.Assimilate()

    msg = "Concluido! {} vistas criadas e paginadas.".format(len(vistas_geradas))
    if erros_callout:
        msg += "\n\nAvisos ({}):\n{}".format(len(erros_callout), "\n".join(erros_callout))
    forms.alert(msg)

if __name__ == '__main__':
    executar_fluxo_callout()