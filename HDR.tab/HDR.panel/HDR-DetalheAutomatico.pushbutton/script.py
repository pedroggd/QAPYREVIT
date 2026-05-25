# -*- coding: utf-8 -*-

__title__ = "Auto\nDetalhe2D"
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
        self.txt_prefixo.Text     = "DET."
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
    import time
    inicio = time.time()    

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

    # ----------------------------------------------------------------
    # Detecta o próximo número disponível para o sufixo atual.
    # Varre todas as vistas existentes que combinem com
    # "prefixo <identificador><N> sufixo" e continua a partir do maior N+1.
    # Se o sufixo mudar (outro pavimento), começa do num_inicial.
    # ----------------------------------------------------------------
    def _proximo_contador(prefixo, identificador, sufixo, zeros, num_inicial):
        pattern = re.compile(
            r'^' + re.escape(prefixo) +
            (r'\s+' if prefixo else r'') +
            re.escape(identificador) +
            r'(\d+)' +
            (r'\s+' + re.escape(sufixo) if sufixo else r'') +
            r'$'
        )
        maior = num_inicial - 1
        for v in FilteredElementCollector(doc).OfClass(DB.View):
            m = pattern.match(v.Name.strip())
            if m:
                try:
                    n = int(m.group(1))
                    if n > maior:
                        maior = n
                except:
                    pass
        return maior + 1

    contador = _proximo_contador(prefixo, identificador, sufixo, zeros, num_inicial)

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

            import math as _math

            mm                  = 1.0 / 304.8
            margem_esq          = 25.0 * mm
            margem_sup          = 15.0 * mm
            margem_inf          = 15.0 * mm
            MARGEM_CORTE        = 5.0  * mm   # Borda de corte da folha (padrão ABNT)
            espacamento         = 10.0 * mm   # Gap minimo horizontal entre viewports
            MARGEM_ENTRE_LINHAS = 20.0 * mm   # Gap minimo vertical entre linhas

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

                doc.Regenerate()

                # Calcula ux_max dinamicamente: viewport do carimbo mais à esquerda - 5mm
                ux_max_calc = bb_sh.Max.X - (175.0 * mm)
                try:
                    vps_carimbo = (FilteredElementCollector(doc, sh.Id)
                                   .OfClass(Viewport)
                                   .ToElements())
                    if vps_carimbo:
                        outlines = []
                        for vp in vps_carimbo:
                            try:
                                ol = vp.GetBoxOutline()
                                outlines.append(ol)
                            except:
                                pass
                        if outlines:
                            outlines.sort(key=lambda o: o.MinimumPoint.X, reverse=True)
                            ux_max_calc = outlines[0].MinimumPoint.X - (5.0 * mm)
                except:
                    pass

                return (
                    sh,
                    bb_sh.Min.X + MARGEM_CORTE + margem_esq,
                    ux_max_calc,
                    bb_sh.Min.Y + MARGEM_CORTE + margem_inf,
                    bb_sh.Max.Y - MARGEM_CORTE - margem_sup,
                )

            sheet, ux_min, ux_max, uy_min, uy_max = nova_prancha()

            # ----------------------------------------------------------------
            # PASSO 1 — Criar todos os viewports em posição de staging e medir
            #           dimensões reais via GetBoxOutline + GetLabelOutline.
            #           slot_w = box_w, slot_h = box_h + label_h
            # ----------------------------------------------------------------
            vp_infos = []

            for info in vistas_geradas:
                v = doc.GetElement(info["id"])
                if v is None:
                    vp_infos.append(None)
                    continue

                if not Viewport.CanAddViewToSheet(doc, sheet.Id, v.Id):
                    erros_callout.append(
                        "Vista '{}' nao pode ser adicionada a prancha.".format(v.Name))
                    vp_infos.append(None)
                    continue

                vp = Viewport.Create(doc, sheet.Id, v.Id, XYZ(0, 0, 0))
                doc.Regenerate()

                # ✅ Fix Detail Number: seta o número correto no bubble do callout
                p_det = v.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                if p_det and not p_det.IsReadOnly:
                    p_det.Set(v.Name.split()[1])
                doc.Regenerate()

                if vp_type_id != ElementId.InvalidElementId:
                    vp.ChangeTypeId(vp_type_id)
                    doc.Regenerate()

                outline  = vp.GetBoxOutline()
                box_w    = outline.MaximumPoint.X - outline.MinimumPoint.X
                box_h    = outline.MaximumPoint.Y - outline.MinimumPoint.Y
                label_h  = 0.0
                try:
                    lbl_outline = vp.GetLabelOutline()
                    if lbl_outline is not None:
                        if lbl_outline.MinimumPoint.Y < outline.MinimumPoint.Y:
                            label_h = lbl_outline.MaximumPoint.Y - lbl_outline.MinimumPoint.Y
                except:
                    pass
                LABEL_H_MIN = 15*mm
                # Mede também a largura do label para incluir no slot_w
                label_w = 0.0
                try:
                    lbl_outline = vp.GetLabelOutline()
                    if lbl_outline is not None:
                        label_w = lbl_outline.MaximumPoint.X - lbl_outline.MinimumPoint.X
                except:
                    pass
                slot_w = max(box_w, label_w)
                vp_infos.append({
                    "vp":      vp,
                    "view_id": v.Id,
                    "box_w":   box_w,
                    "box_h":   box_h,
                    "label_h": label_h,
                    "label_w": label_w,
                    "slot_w":  slot_w,
                    "slot_h":  box_h + max(label_h, LABEL_H_MIN),
                    "nome":    v.Name,
                })

            # Filtra Nones
            vp_infos = [vi for vi in vp_infos if vi is not None]

            if not vp_infos:
                t.Commit()
                tg.Assimilate()
                forms.alert("Nenhum viewport pôde ser criado.")
                return

            # ----------------------------------------------------------------
            # PASSO 2 — Grade ótima: calcula ncols/nrows com base no maior
            #           viewport. Distribui igualmente entre pranchas.
            # ----------------------------------------------------------------
            # ----------------------------------------------------------------
            # PASSO 2 — Bin packing: vai adicionando viewports e abre nova
            #           prancha quando não couber mais.
            # ----------------------------------------------------------------
            area_w = ux_max - ux_min
            area_h = uy_max - uy_min

            grupo_atual = []
            grupos_prancha = []

            for vi in vp_infos:
                grupo_atual.append(vi)

                ref_sw_p = max(v["slot_w"] for v in grupo_atual) 
                ref_sh_p = max(v["slot_h"] for v in grupo_atual) 
                ncols_t = max(1, int((area_w + espacamento) / (ref_sw_p + espacamento)))
                nrows_t = max(1, int((area_h + MARGEM_ENTRE_LINHAS) / (ref_sh_p + MARGEM_ENTRE_LINHAS)))

                if len(grupo_atual) > ncols_t * nrows_t:
                    grupos_prancha.append(grupo_atual[:-1])
                    grupo_atual = [vi]

            if grupo_atual:
                grupos_prancha.append(grupo_atual)

            # ----------------------------------------------------------------
            # PASSO 3 — Posicionamento com espaçamento uniforme e centralizado.
            #           LabelOffset recalculado após o move.
            #           Ao mudar de prancha, recria viewports corretamente.
            # ----------------------------------------------------------------
            def _recriar_vp_na_prancha(sh_dest, vi):
                """Deleta viewport existente e recria na nova prancha."""
                try:
                    doc.Delete(vi["vp"].Id)
                    doc.Regenerate()
                except:
                    pass

                new_vp = Viewport.Create(doc, sh_dest.Id, vi["view_id"], XYZ(0, 0, 0))
                doc.Regenerate()

                # ✅ Fix Detail Number: seta o número correto no bubble do callout
                vista_elem = doc.GetElement(vi["view_id"])
                if vista_elem:
                    p_det = vista_elem.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                    if p_det and not p_det.IsReadOnly:
                        p_det.Set(vi["nome"].split()[1])
                doc.Regenerate()

                if new_vp and vp_type_id != ElementId.InvalidElementId:
                    new_vp.ChangeTypeId(vp_type_id)
                    doc.Regenerate()

                if new_vp:
                    ol    = new_vp.GetBoxOutline()
                    bw    = ol.MaximumPoint.X - ol.MinimumPoint.X
                    bh    = ol.MaximumPoint.Y - ol.MinimumPoint.Y
                    lh    = 0.0
                    lw    = 0.0
                    try:
                        lbl_ol = new_vp.GetLabelOutline()
                        if lbl_ol is not None and lbl_ol.MinimumPoint.Y < ol.MinimumPoint.Y:
                            lh = lbl_ol.MaximumPoint.Y - lbl_ol.MinimumPoint.Y
                        if lbl_ol is not None:
                            lw = lbl_ol.MaximumPoint.X - lbl_ol.MinimumPoint.X
                    except:
                        pass
                    vi["vp"]      = new_vp
                    vi["box_w"]   = bw
                    vi["box_h"]   = bh
                    vi["label_h"] = lh
                    vi["label_w"] = lw
                    vi["slot_w"]  = max(bw, lw)
                    vi["slot_h"]  = bh + max(lh, LABEL_H_MIN)

            primeira_prancha = sheet

            for p_idx, grupo in enumerate(grupos_prancha):
                if p_idx == 0:
                    sh_atual = primeira_prancha
                    ux_min_p, ux_max_p, uy_min_p, uy_max_p = ux_min, ux_max, uy_min, uy_max
                else:
                    sh_atual, ux_min_p, ux_max_p, uy_min_p, uy_max_p = nova_prancha()
                    for vi in grupo:
                        _recriar_vp_na_prancha(sh_atual, vi)

                area_w_p = ux_max_p - ux_min_p
                area_h_p = uy_max_p - uy_min_p

                # Divide o grupo em linhas de ncols
                # Recalcula ncols pro grupo desta prancha
                ref_sw_p = sum(vi["slot_w"] for vi in grupo) / len(grupo)
                ncols = max(1, int((area_w_p + espacamento) / (ref_sw_p + espacamento)))

                linhas_g = [grupo[i:i + ncols] for i in range(0, len(grupo), ncols)]
                n_lin    = len(linhas_g)

                alt_lins = [max(vi["slot_h"] for vi in ln) for ln in linhas_g]
                tot_h    = sum(alt_lins)

                # gap_y uniforme, nunca menor que MARGEM_ENTRE_LINHAS
                GAP_Y_MAX = 40.0 * mm
                if n_lin > 1:
                    gap_y = min(GAP_Y_MAX, max(MARGEM_ENTRE_LINHAS,
                                (area_h_p - tot_h) / float(n_lin - 1)))
                else:
                    gap_y = 0.0

                # Começa do topo da área útil (sem centralização vertical)
                cur_y_p = uy_max_p

                for l_idx, linha in enumerate(linhas_g):
                    rh_max = alt_lins[l_idx]
                    n_col  = len(linha)
                    tot_w  = sum(vi["slot_w"] for vi in linha)

                    # gap_x uniforme, nunca menor que espacamento
                    GAP_X_MAX = 20.0 * mm
                    if n_col > 1:
                        gap_x = min(GAP_X_MAX, max(espacamento,
                                    (area_w_p - tot_w) / float(n_col - 1)))
                    else:
                        gap_x = 0.0

                    # Centraliza horizontalmente
                    bloco_w = tot_w + gap_x * (n_col - 1)
                    cur_x_p = ux_min_p + (area_w_p - bloco_w) / 2.0
                    base_y = cur_y_p - rh_max

                    for vi in linha:
                        try:
                            vp = vi["vp"]
                            if vp is None:
                                continue
                            ol = vp.GetBoxOutline()
                            # Alinha todos pela borda inferior do box
                            DB.ElementTransformUtils.MoveElement(
                                doc, vp.Id,
                                XYZ(cur_x_p - ol.MinimumPoint.X,
                                    base_y  - ol.MinimumPoint.Y, 0))
                            # Zera o LabelOffset — Revit posiciona o título
                            # sempre abaixo do box, igual pra qualquer tamanho
                            GAP = 9*mm
                            vp.LabelOffset = XYZ(0.0, -GAP, 0.0)
                            doc.Regenerate()

                        except Exception as e:
                            erros_callout.append(
                                "Erro ao posicionar vista '{}': {}".format(vi["nome"], str(e)))

                        cur_x_p += vi["slot_w"] + gap_x

                    cur_y_p -= (rh_max + gap_y)

            t.Commit()
        tg.Assimilate()

        tempo = time.time() - inicio
    mins = int(tempo // 60)
    segs = int(tempo % 60)
    msg = "Concluido! {} vistas criadas e paginadas.".format(len(vistas_geradas))
    msg += "\n\nTempo de execução: {}m {}s".format(mins, segs)
    if erros_callout:
        msg += "\n\nAvisos ({}):\n{}".format(len(erros_callout), "\n".join(erros_callout))
    forms.alert(msg)

if __name__ == '__main__':
    executar_fluxo_callout()