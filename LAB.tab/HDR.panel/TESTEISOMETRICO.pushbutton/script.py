# -*- coding: utf-8 -*-

__title__ = "AutoIso3D\n+"
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
    FilteredElementCollector, FamilySymbol, ViewFamilyType, ViewFamily,
    ViewSheet, Transaction, TransactionGroup, BuiltInCategory, BuiltInParameter,
    Viewport, XYZ, View3D, BoundingBoxXYZ, ElementId, StorageType,
    ViewType, Level, Options, Line, GeometryInstance
)
import Autodesk.Revit.DB as DB
import Autodesk.Revit.Exceptions as Exceptions
from Autodesk.Revit.UI.Selection import PickBoxStyle

logger = script.get_logger()
doc   = revit.doc
uidoc = revit.uidoc

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
        u'\u00c1': u'A', u'\u00c0': u'A', u'\u00c3': u'A', u'\u00c2': u'A',
        u'\u00c9': u'E', u'\u00c8': u'E', u'\u00ca': u'E',
        u'\u00cd': u'I', u'\u00cc': u'I', u'\u00ce': u'I',
        u'\u00d3': u'O', u'\u00d2': u'O', u'\u00d5': u'O', u'\u00d4': u'O',
        u'\u00da': u'U', u'\u00d9': u'U', u'\u00db': u'U',
        u'\u00c7': u'C',
        u'\u00e1': u'a', u'\u00e0': u'a', u'\u00e3': u'a', u'\u00e2': u'a',
        u'\u00e9': u'e', u'\u00e8': u'e', u'\u00ea': u'e',
        u'\u00ed': u'i', u'\u00ec': u'i', u'\u00ee': u'i',
        u'\u00f3': u'o', u'\u00f2': u'o', u'\u00f5': u'o', u'\u00f4': u'o',
        u'\u00fa': u'u', u'\u00f9': u'u', u'\u00fb': u'u',
        u'\u00e7': u'c', u'\u00ba': u'o', u'\u00aa': u'a',
    }
    for k, v in mapping.items():
        s = s.replace(k, v)
    return s

def sanitize_name(name):
    return re.sub(r'[\\:\{\}\[\]|;<>?\'~]', '-', name).strip()

def _proximo_contador(prefixo, identificador, num_inicial, sufixo):
    maior_num = num_inicial - 1
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate: continue
        p_pref  = re.escape(prefixo)       if prefixo       else ""
        p_ident = re.escape(identificador) if identificador else ""
        p_suf   = re.escape(sufixo)        if sufixo        else ""
        pattern = r'^'
        if p_pref:  pattern += p_pref  + r'\s*'
        if p_ident: pattern += p_ident
        pattern += r'(\d+)'
        if p_suf:   pattern += r'\s*' + p_suf
        pattern += r'$'
        match = re.search(pattern, v.Name, re.IGNORECASE)
        if match:
            try:
                num = int(match.group(1))
                if num > maior_num:
                    maior_num = num
            except:
                pass
    return maior_num + 1

def get_unique_view_name(base_name):
    existing = {v.Name for v in FilteredElementCollector(doc).OfClass(DB.View)}
    if base_name not in existing:
        return base_name
    i = 1
    while True:
        candidate = "%s (%d)" % (base_name, i)
        if candidate not in existing:
            return candidate
        i += 1

def get_titleblock_symbols_by_family():
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

def get_3d_view_family_types():
    tipos = {}
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.ThreeDimensional:
            p = vft.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            nome = p.AsString() if p else vft.Id.ToString()
            tipos[nome] = vft
    return tipos

def get_viewport_types():
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

def criar_placeholder_drafting_view(nome_base, drafting_vft_id):
    nome = get_unique_view_name(nome_base + " - REF")
    nova_drafting = DB.ViewDrafting.Create(doc, drafting_vft_id)
    nova_drafting.Name = nome
    try:
        pt1 = XYZ(0, 0, 0)
        pt2 = XYZ(0.01, 0, 0)
        linha = DB.Line.CreateBound(pt1, pt2)
        doc.Create.NewDetailCurve(nova_drafting, linha)
    except:
        pass
    return nova_drafting


# ---------------------------------------------------------------------------
# JANELA DE CONFIGURAÇÃO
# ---------------------------------------------------------------------------
class IsometricConfigWindow(Window):
    COR_AZUL   = Color.FromRgb(30, 80, 160)
    COR_VERDE  = Color.FromRgb(22, 160, 80)
    COR_CINZA  = Color.FromRgb(120, 120, 120)
    COR_BRANCO = Color.FromRgb(255, 255, 255)
    COR_FUNDO  = Color.FromRgb(245, 245, 245)

    def __init__(self, templates_dict, view_types_dict, titleblock_symbols, viewport_types_dict):
        self.templates_dict      = templates_dict
        self.view_types_dict     = view_types_dict
        self.viewport_types_dict = viewport_types_dict
        self.resultado           = None

        self._template_names  = sorted(templates_dict.keys())
        self._view_type_names = sorted(view_types_dict.keys())
        self.campos_carimbo   = {}

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

        self._tb_labels  = []
        self._tb_symbols = {}

        def _add_family(fam_name):
            if fam_name not in titleblock_symbols:
                return
            for tipo in sorted(titleblock_symbols[fam_name].keys()):
                lbl = "%s : %s" % (fam_name, tipo)
                self._tb_labels.append(lbl)
                self._tb_symbols[lbl] = titleblock_symbols[fam_name][tipo]

        _add_family(FAMILIA_NOME)
        for fam in sorted(titleblock_symbols.keys()):
            if fam != FAMILIA_NOME:
                _add_family(fam)

        self.Title         = "Configuracao de Isometricos"
        self.Width         = 500
        self.SizeToContent = System.Windows.SizeToContent.Height
        self.ResizeMode    = ResizeMode.CanResizeWithGrip
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background            = SolidColorBrush(self.COR_FUNDO)

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)
        root.RowDefinitions.Add(self._row(GridLength.Auto))
        root.RowDefinitions.Add(self._row(GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(self._row(GridLength.Auto))

        header = StackPanel()
        self._lbl(header, "Criar Isometricos+", 16, bold=True, cor=self.COR_AZUL, mg=(0, 0, 0, 2))
        header.Children.Add(self._sep(10))
        Grid.SetRow(header, 0)
        root.Children.Add(header)

        self.tabs = TabControl()
        self.tabs.FontSize = 12
        self.tabs.Margin   = Thickness(0, 0, 0, 12)
        Grid.SetRow(self.tabs, 1)
        root.Children.Add(self.tabs)

        self._build_tab_config()
        self._build_tab_prancha()
        self._build_tab_carimbo()

        btn_ok = Button()
        btn_ok.Content    = "OK - Iniciar Selecao de Area"
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
        if mg:   lbl.Margin     = Thickness(*mg)
        if parent is not None:
            parent.Children.Add(lbl)
        return lbl

    def _build_tab_config(self):
        tab = TabItem()
        tab.Header = "Configuracoes"
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        pnl = StackPanel()
        pnl.Margin = Thickness(12)

        self._lbl(pnl, "Nome / Prefixo (ex: 'DETALHE', 'ISO'):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_prefixo = TextBox()
        self.txt_prefixo.FontSize = 12
        self.txt_prefixo.Padding  = Thickness(6, 4, 6, 4)
        self.txt_prefixo.Margin   = Thickness(0, 0, 0, 12)
        self.txt_prefixo.Text     = "DET."
        pnl.Children.Add(self.txt_prefixo)

        self._lbl(pnl, "Identificador (ex: 'H', 'S', ou vazio):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_identificador = TextBox()
        self.txt_identificador.FontSize = 12
        self.txt_identificador.Padding  = Thickness(6, 4, 6, 4)
        self.txt_identificador.Margin   = Thickness(0, 0, 0, 12)
        self.txt_identificador.Text     = "H"
        pnl.Children.Add(self.txt_identificador)

        g_cont = Grid()
        g_cont.Margin = Thickness(0, 0, 0, 12)
        g_cont.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g_cont.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))

        pnl_num = StackPanel(); pnl_num.Margin = Thickness(0, 0, 4, 0)
        self._lbl(pnl_num, "Numero Inicial:", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_num_inicial = TextBox()
        self.txt_num_inicial.FontSize = 12
        self.txt_num_inicial.Padding  = Thickness(6, 4, 6, 4)
        self.txt_num_inicial.Text     = "01"
        pnl_num.Children.Add(self.txt_num_inicial)
        Grid.SetColumn(pnl_num, 0); g_cont.Children.Add(pnl_num)

        pnl_zero = StackPanel(); pnl_zero.Margin = Thickness(4, 0, 0, 0)
        self._lbl(pnl_zero, "Formato do Numero:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_zeros = ComboBox(); self.combo_zeros.FontSize = 11
        for label in ["1 digito (1, 2, 3)", "2 digitos (01, 02)", "3 digitos (001, 002)"]:
            it = ComboBoxItem(); it.Content = label
            self.combo_zeros.Items.Add(it)
        self.combo_zeros.SelectedIndex = 1
        pnl_zero.Children.Add(self.combo_zeros)
        Grid.SetColumn(pnl_zero, 1); g_cont.Children.Add(pnl_zero)

        pnl.Children.Add(g_cont)

        self._lbl(pnl, "Sufixo (ex: '- 1 TIPO' ou vazio):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_sufixo = TextBox()
        self.txt_sufixo.FontSize = 12
        self.txt_sufixo.Padding  = Thickness(6, 4, 6, 4)
        self.txt_sufixo.Margin   = Thickness(0, 0, 0, 16)
        self.txt_sufixo.Text     = ""
        pnl.Children.Add(self.txt_sufixo)

        pnl.Children.Add(self._sep(10))

        self._lbl(pnl, "Template de Vista 3D:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_template = ComboBox()
        self.combo_template.FontSize = 11
        self.combo_template.Margin   = Thickness(0, 0, 0, 12)
        for nome in self._template_names:
            item = ComboBoxItem(); item.Content = nome
            self.combo_template.Items.Add(item)
        self.combo_template.SelectedIndex = 0
        pnl.Children.Add(self.combo_template)

        self._lbl(pnl, "Tipo de Vista 3D:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_view_type = ComboBox()
        self.combo_view_type.FontSize = 11
        self.combo_view_type.Margin   = Thickness(0, 0, 0, 16)
        for nome in self._view_type_names:
            item = ComboBoxItem(); item.Content = nome
            self.combo_view_type.Items.Add(item)
        self.combo_view_type.SelectedIndex = 0
        pnl.Children.Add(self.combo_view_type)

        pnl.Children.Add(self._sep(10))

        self._lbl(pnl, "Escala (ex: 25 = 1:25):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_escala = TextBox()
        self.txt_escala.FontSize = 12
        self.txt_escala.Padding  = Thickness(6, 4, 6, 4)
        self.txt_escala.Margin   = Thickness(0, 0, 0, 12)
        self.txt_escala.Text     = "25"
        pnl.Children.Add(self.txt_escala)

        self.chk_sem_escala = CheckBox()
        self.chk_sem_escala.Content   = "Sem Escala (ignora o campo acima)"
        self.chk_sem_escala.FontSize  = 11
        self.chk_sem_escala.Margin    = Thickness(0, 0, 0, 12)
        self.chk_sem_escala.IsChecked = True
        pnl.Children.Add(self.chk_sem_escala)

        scroll.Content = pnl
        tab.Content    = scroll
        self.tabs.Items.Add(tab)

    def _build_tab_prancha(self):
        tab = TabItem()
        tab.Header = "Prancha"
        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        pnl = StackPanel()
        pnl.Margin = Thickness(12)

        self._lbl(pnl, "Tipo de Folha (Carimbo):", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_titleblock = ComboBox()
        self.combo_titleblock.FontSize = 11
        self.combo_titleblock.Margin   = Thickness(0, 0, 0, 16)
        for lbl in self._tb_labels:
            item = ComboBoxItem(); item.Content = lbl
            self.combo_titleblock.Items.Add(item)
        self.combo_titleblock.SelectedIndex = 0
        pnl.Children.Add(self.combo_titleblock)

        pnl.Children.Add(self._sep(8))

        self._lbl(pnl, "Tipo de Viewport:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_vp_type = ComboBox()
        self.combo_vp_type.FontSize = 11
        self.combo_vp_type.Margin   = Thickness(0, 0, 0, 16)
        for lbl in self._vp_type_labels:
            item = ComboBoxItem(); item.Content = lbl
            self.combo_vp_type.Items.Add(item)
        self.combo_vp_type.SelectedIndex = self._vp_type_best
        pnl.Children.Add(self.combo_vp_type)

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

        self._lbl(outer, "Parametros do carimbo", 12, bold=True, mg=(0, 0, 0, 4))
        self._lbl(outer, "Aplicados em TODAS as pranchas criadas.", 10,
                  cor=self.COR_CINZA, mg=(0, 0, 0, 10))
        outer.Children.Add(self._sep(8))

        for param_name, label_text in PARAMS_CARIMBO:
            g = Grid()
            g.Margin = Thickness(0, 0, 0, 6)
            g.ColumnDefinitions.Add(self._col(GridLength(150)))
            g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))

            lbl = self._lbl(None, label_text, 11, mg=(0, 4, 8, 0))
            Grid.SetColumn(lbl, 0); g.Children.Add(lbl)

            txt = TextBox()
            txt.FontSize = 11
            txt.Padding  = Thickness(6, 4, 6, 4)
            Grid.SetColumn(txt, 1); g.Children.Add(txt)

            outer.Children.Add(g)
            self.campos_carimbo[label_text] = txt

        scroll.Content = outer
        tab.Content    = scroll
        self.tabs.Items.Add(tab)

    def on_ok(self, sender, args):
        prefixo       = self.txt_prefixo.Text.strip()
        identificador = self.txt_identificador.Text.strip()
        sufixo        = self.txt_sufixo.Text.strip()

        try:
            num_inicial = int(self.txt_num_inicial.Text.strip())
        except:
            num_inicial = 1

        zeros = self.combo_zeros.SelectedIndex + 1

        t_idx       = self.combo_template.SelectedIndex
        template_id = (self.templates_dict[self._template_names[t_idx]]
                       if 0 <= t_idx < len(self._template_names)
                       else ElementId.InvalidElementId)

        vt_idx       = self.combo_view_type.SelectedIndex
        view_type_id = (self.view_types_dict[self._view_type_names[vt_idx]].Id
                        if 0 <= vt_idx < len(self._view_type_names)
                        else ElementId.InvalidElementId)

        try:
            escala = int(self.txt_escala.Text.strip())
        except:
            escala = 25

        sem_escala = bool(self.chk_sem_escala.IsChecked)

        tb_idx    = self.combo_titleblock.SelectedIndex
        tb_lbl    = self._tb_labels[tb_idx] if 0 <= tb_idx < len(self._tb_labels) else None
        tb_symbol = self._tb_symbols.get(tb_lbl) if tb_lbl else None

        vp_idx         = self.combo_vp_type.SelectedIndex
        vp_lbl         = self._vp_type_labels[vp_idx] if 0 <= vp_idx < len(self._vp_type_labels) else None
        vp_type_id_sel = self.viewport_types_dict.get(vp_lbl, ElementId.InvalidElementId) if vp_lbl else ElementId.InvalidElementId

        dados_carimbo = {}
        for label_text, txt in self.campos_carimbo.items():
            valor = txt.Text.strip()
            if valor:
                dados_carimbo[label_text] = valor

        self.resultado = {
            "prefixo":       prefixo,
            "identificador": identificador,
            "num_inicial":   num_inicial,
            "zeros":         zeros,
            "sufixo":        sufixo,
            "template_id":   template_id,
            "view_type_id":  view_type_id,
            "escala":        escala,
            "sem_escala":    sem_escala,
            "tb_symbol":     tb_symbol,
            "vp_type_id":    vp_type_id_sel,
            "dados_carimbo": dados_carimbo,
        }
        self.Close()


# ---------------------------------------------------------------------------
# FLUXO PRINCIPAL
# ---------------------------------------------------------------------------
def executar_fluxo_isometrico():
    import math as _math

    titleblock_symbols = get_titleblock_symbols_by_family()
    if not titleblock_symbols:
        forms.alert(
            "Nenhum carimbo encontrado.\nCarregue a familia '%s' e tente novamente." % FAMILIA_NOME,
            exitscript=True
        )

    templates = {"(Nenhum)": ElementId.InvalidElementId}
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate and v.ViewType == ViewType.ThreeD:
            templates[v.Name] = v.Id
    if len(templates) == 1:
        for v in FilteredElementCollector(doc).OfClass(DB.View):
            if v.IsTemplate:
                templates[v.Name] = v.Id

    view_types_3d = get_3d_view_family_types()
    if not view_types_3d:
        forms.alert("Nenhum tipo de vista 3D encontrado no projeto.", exitscript=True)

    viewport_types_dict = get_viewport_types()

    dlg = IsometricConfigWindow(templates, view_types_3d, titleblock_symbols, viewport_types_dict)
    dlg.ShowDialog()

    if not dlg.resultado:
        script.exit()

    prefixo       = dlg.resultado["prefixo"]
    identificador = dlg.resultado["identificador"]
    num_inicial   = dlg.resultado["num_inicial"]
    zeros         = dlg.resultado["zeros"]
    sufixo        = dlg.resultado["sufixo"]
    template_id   = dlg.resultado["template_id"]
    view_type_id  = dlg.resultado["view_type_id"]
    escala        = dlg.resultado["escala"]
    sem_escala    = dlg.resultado["sem_escala"]
    tb_symbol     = dlg.resultado["tb_symbol"]
    vp_type_id    = dlg.resultado["vp_type_id"]
    dados_carimbo = dlg.resultado["dados_carimbo"]

    if not tb_symbol:
        forms.alert("Nenhum tipo de folha selecionado.", exitscript=True)
    tb_type_id = tb_symbol.Id

    vista_ativa  = doc.ActiveView
    nivel_atual  = vista_ativa.GenLevel

    drafting_vft_id = ElementId.InvalidElementId
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.Drafting:
            drafting_vft_id = vft.Id
            break

    # -----------------------------------------------------------------------
    # CONSTANTES DE LAYOUT (definidas aqui pois sao usadas em nova_prancha
    # e em todo o loop de paginacao)
    # -----------------------------------------------------------------------
    mm                  = 1.0 / 304.8
    margem_esq          = 10.0 * mm
    margem_sup          = 10.0 * mm
    margem_inf          = 10.0 * mm
    MARGEM_CORTE        = 1.0  * mm
    espacamento         = 2.0  * mm
    MARGEM_ENTRE_LINHAS = 5.0  * mm
    LABEL_H_MIN         = 8.0  * mm
    GAP_LABEL           = 1.5  * mm

    # -----------------------------------------------------------------------
    # HELPERS de transacao atomica
    # -----------------------------------------------------------------------
    def _exec_transaction(nome, func):
        """Executa func() dentro de uma transacao propria. Retorna o resultado."""
        with Transaction(doc, nome) as t:
            t.Start()
            try:
                resultado = func()
                t.Commit()
                return resultado
            except Exception as e:
                t.RollBack()
                raise e



    def detectar_max_por_prancha(bb):
        w_mm = (bb.Max.X - bb.Min.X) * 304.8
        h_mm = (bb.Max.Y - bb.Min.Y) * 304.8
        paisagem = w_mm > h_mm
        area = w_mm * h_mm
        if area > 900000:
            return 12 if paisagem else 15
        elif area > 450000:
            return 6 if paisagem else 6
        else:
            return 3


    def nova_prancha():
        """Cria uma nova folha e retorna (sheet, ux_min, ux_max, uy_min, uy_max).
        DEVE ser chamada dentro de uma transacao ativa."""
        sh = ViewSheet.Create(doc, tb_type_id)

    
        doc.Regenerate()

        tbs = (FilteredElementCollector(doc, sh.Id)
               .OfCategory(BuiltInCategory.OST_TitleBlocks)
               .WhereElementIsNotElementType().ToElements())
        tb_instance = tbs[0] if tbs else None

        if not tb_instance:
            raise Exception("Prancha sem carimbo.")

        for param_name, valor in dados_carimbo.items():
            p = sh.LookupParameter(param_name)
            if not p:
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

        bb_sh = tb_instance.get_BoundingBox(sh)
        if not bb_sh:
            raise Exception("BoundingBox None.")

        doc.Regenerate()

        ux_max_calc = bb_sh.Max.X - (175.0 * mm)
        try:
            vps_carimbo = (FilteredElementCollector(doc, sh.Id)
                           .OfClass(Viewport).ToElements())
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
            detectar_max_por_prancha(bb_sh),
        )

    # -----------------------------------------------------------------------
    # FASE 1 — SELECAO INTERATIVA DE AREAS
    # -----------------------------------------------------------------------
    grupos              = []
    marcacoes_por_caixa = []
    contador            = _proximo_contador(prefixo, identificador, num_inicial, sufixo)
    contador_inicial    = contador

    def _limpar_marcacoes(grupos_marc=None):
        from System.Collections.Generic import List
        ids_a_deletar = List[DB.ElementId]()
        if grupos_marc is None:
            for g in marcacoes_por_caixa:
                for eid in g:
                    if doc.GetElement(eid) is not None:
                        ids_a_deletar.Add(eid)
            marcacoes_por_caixa[:] = []
        else:
            for g in grupos_marc:
                for eid in g:
                    if doc.GetElement(eid) is not None:
                        ids_a_deletar.Add(eid)
        if ids_a_deletar.Count == 0:
            return
        try:
            with Transaction(doc, "Limpar Marcacoes Temporarias") as t_limpa:
                t_limpa.Start()
                doc.Delete(ids_a_deletar)
                t_limpa.Commit()
        except Exception:
            pass

    try:
        while True:
            numero_str    = str(contador).zfill(zeros)
            contador_full = "%s%s" % (identificador, numero_str)
            partes        = [prefixo, contador_full, sufixo]
            nome_preview  = " ".join([p for p in partes if p])

            instrucao = "[%s]  (%d ja marcado(s))  - Desenhe o retangulo ou ESC para opcoes" % (
                nome_preview, len(grupos)
            )

            try:
                box = uidoc.Selection.PickBox(PickBoxStyle.Directional, instrucao)
                grupos.append(box)

                ids_marcacao = []
                try:
                    with Transaction(doc, "Marcacao: %s" % nome_preview) as t_marc:
                        t_marc.Start()
                        min_x = min(box.Min.X, box.Max.X)
                        max_x = max(box.Min.X, box.Max.X)
                        min_y = min(box.Min.Y, box.Max.Y)
                        max_y = max(box.Min.Y, box.Max.Y)
                        for p1, p2 in [
                            (XYZ(min_x, min_y, 0), XYZ(max_x, min_y, 0)),
                            (XYZ(max_x, min_y, 0), XYZ(max_x, max_y, 0)),
                            (XYZ(max_x, max_y, 0), XYZ(min_x, max_y, 0)),
                            (XYZ(min_x, max_y, 0), XYZ(min_x, min_y, 0)),
                        ]:
                            dl = doc.Create.NewDetailCurve(vista_ativa, Line.CreateBound(p1, p2))
                            ids_marcacao.append(dl.Id)
                        centro = XYZ((min_x + max_x) / 2.0, (min_y + max_y) / 2.0, 0)
                        txt_type = (DB.FilteredElementCollector(doc)
                                    .OfClass(DB.TextNoteType).FirstElement())
                        if txt_type is not None:
                            opts = DB.TextNoteOptions(txt_type.Id)
                            opts.HorizontalAlignment = DB.HorizontalTextAlignment.Center
                            tnote = DB.TextNote.Create(
                                doc, vista_ativa.Id, centro, nome_preview, opts)
                            ids_marcacao.append(tnote.Id)
                        t_marc.Commit()
                except Exception:
                    pass

                marcacoes_por_caixa.append(ids_marcacao)
                contador += 1

            except Exceptions.OperationCanceledException:
                if grupos:
                    opcao = forms.alert(
                        "Selecao pausada. %d isometrico(s) na fila.\nO que deseja fazer?" % len(grupos),
                        options=["Finalizar e Criar Pranchas", "Desfazer o ULTIMO e Continuar", "Cancelar Script"]
                    )
                    if opcao == "Desfazer o ULTIMO e Continuar":
                        grupos.pop()
                        contador -= 1
                        if marcacoes_por_caixa:
                            _limpar_marcacoes(grupos_marc=[marcacoes_por_caixa.pop()])
                        continue
                    elif opcao == "Finalizar e Criar Pranchas":
                        _limpar_marcacoes()
                        break
                    else:
                        _limpar_marcacoes()
                        script.exit()
                else:
                    script.exit()
            except Exception as e:
                _limpar_marcacoes()
                forms.alert("Erro na selecao: %s" % str(e))
                break
    finally:
        _limpar_marcacoes()

    if not grupos:
        forms.alert("Nenhum isometrico selecionado. Operacao cancelada.")
        script.exit()

    # -----------------------------------------------------------------------
    # FASE 2 — CRIACAO DAS VISTAS 3D (cada uma em transacao propria,
    #           todas dentro do TransactionGroup)
    # -----------------------------------------------------------------------
    vistas_geradas = []
    erros          = []

    with TransactionGroup(doc, "Gerador de Isometricos") as tg:
        tg.Start()

        # --- 2a. Cria as vistas isometricas ---
        for idx, box in enumerate(grupos):
            numero_str    = str(contador_inicial + idx).zfill(zeros)
            contador_full = "%s%s" % (identificador, numero_str)
            partes        = [prefixo, contador_full, sufixo]
            nome_iso      = sanitize_name(" ".join([p for p in partes if p]))

            novo_id        = None
            placeholder_id = None

            with Transaction(doc, "Criar Isometrico: %s" % nome_iso) as t:
                t.Start()
                try:
                    p_min_x = min(box.Min.X, box.Max.X)
                    p_max_x = max(box.Min.X, box.Max.X)
                    p_min_y = min(box.Min.Y, box.Max.Y)
                    p_max_y = max(box.Min.Y, box.Max.Y)

                    margem = 0.5 / 0.3048

                    if nivel_atual:
                        z_level = nivel_atual.Elevation
                        z_min   = z_level - margem
                        z_max   = z_level + (2.30 / 0.3048)
                    else:
                        z_base = min(box.Min.Z, box.Max.Z)
                        z_min  = z_base - margem
                        z_max  = z_base + (2.30 / 0.3048)

                    section_box     = DB.BoundingBoxXYZ()
                    section_box.Min = XYZ(p_min_x - margem, p_min_y - margem, z_min)
                    section_box.Max = XYZ(p_max_x + margem, p_max_y + margem, z_max)

                    nova      = View3D.CreateIsometric(doc, view_type_id)
                    nova.Name = get_unique_view_name(nome_iso)
                    nova.IsSectionBoxActive = True
                    nova.SetSectionBox(section_box)

                    if template_id != ElementId.InvalidElementId:
                        try:
                            nova.ViewTemplateId = template_id
                            doc.Regenerate()
                        except:
                            pass

                    if not sem_escala:
                        try:
                            nova.Scale = escala
                        except:
                            pass

                    if (drafting_vft_id != ElementId.InvalidElementId
                            and vista_ativa.ViewType in [
                                ViewType.FloorPlan,
                                ViewType.EngineeringPlan,
                                ViewType.CeilingPlan]):
                        try:
                            placeholder    = criar_placeholder_drafting_view(nome_iso, drafting_vft_id)
                            placeholder_id = placeholder.Id
                            pt1 = XYZ(p_min_x, p_min_y, box.Min.Z)
                            pt2 = XYZ(p_max_x, p_max_y, box.Min.Z)
                            DB.ViewSection.CreateReferenceCallout(
                                doc, vista_ativa.Id, placeholder.Id, pt1, pt2)
                        except Exception as e_ref:
                            erros.append("Aviso: Falha ao criar referencia para '%s': %s"
                                         % (nome_iso, str(e_ref)))

                    doc.Regenerate()
                    novo_id = nova.Id
                    t.Commit()

                except Exception as e:
                    erros.append("Erro ao criar '%s': %s" % (nome_iso, str(e)))
                    t.RollBack()

            if novo_id is not None:
                vistas_geradas.append({
                    "id":             novo_id,
                    "placeholder_id": placeholder_id,
                    "numero_detalhe": contador_inicial + idx,
                })

        if not vistas_geradas:
            tg.RollBack()
            forms.alert("Nenhuma vista criada.\n" + "\n".join(erros))
            script.exit()

        # --- 2b. Cria a primeira prancha e adiciona todos os viewports ---
        sheet_result = [None]  # usa lista para ser mutavel dentro de closures
        sheet_bounds = [None]  # (ux_min, ux_max, uy_min, uy_max)

        def _criar_primeira_prancha():
            resultado = nova_prancha()
            sheet_result[0] = resultado[0]
            sheet_bounds[0] = resultado[1:]
            return resultado

        with Transaction(doc, "Criar Primeira Prancha") as t:
            t.Start()
            try:
                sh, ux_min, ux_max, uy_min, uy_max, max_por_prancha = nova_prancha()
                t.Commit()
            except Exception as e:
                t.RollBack()
                tg.RollBack()
                forms.alert("Erro ao criar prancha: %s" % str(e))
                script.exit()

        # --- 2c. Adiciona viewports na primeira prancha ---
        vp_infos = []

        with Transaction(doc, "Adicionar Viewports na Prancha") as t:
            t.Start()
            try:
                doc.Regenerate()

                for info in vistas_geradas:
                    v_3d = doc.GetElement(info["id"])
                    if v_3d is None:
                        continue

                    if not Viewport.CanAddViewToSheet(doc, sh.Id, v_3d.Id):
                        erros.append("Vista '%s' nao pode ser adicionada a prancha." % v_3d.Name)
                        continue

                    placeholder_id = info.get("placeholder_id")
                    vp_3d          = None
                    vp_fantasma_id = None

                    if placeholder_id and Viewport.CanAddViewToSheet(doc, sh.Id, placeholder_id):
                        try:
                            vp_fantasma = Viewport.Create(doc, sh.Id, placeholder_id, XYZ(10, 10, 0))
                            doc.Regenerate()
                            if vp_fantasma:
                                vp_fantasma_id = vp_fantasma.Id
                                vista_fantasma = doc.GetElement(placeholder_id)
                                if vista_fantasma:
                                    p_det_f = vista_fantasma.get_Parameter(
                                        BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                    if p_det_f and not p_det_f.IsReadOnly:
                                        parts = vista_fantasma.Name.split()
                                        if len(parts) > 1:
                                            p_det_f.Set(parts[1])

                            vp_3d = Viewport.Create(doc, sh.Id, v_3d.Id, XYZ(0, 0, 0))
                            doc.Regenerate()
                            if vp_3d:
                                p_3d = v_3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                if p_3d and not p_3d.IsReadOnly:
                                    parts = v_3d.Name.split()
                                    if len(parts) > 1:
                                        p_3d.Set(parts[1] + u"\u200b")
                        except Exception as e_vp:
                            erros.append("Aviso VP fantasma '%s': %s" % (v_3d.Name, str(e_vp)))
                    else:
                        try:
                            vp_3d = Viewport.Create(doc, sh.Id, v_3d.Id, XYZ(0, 0, 0))
                            doc.Regenerate()
                            if vp_3d:
                                p_3d = v_3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                if p_3d and not p_3d.IsReadOnly:
                                    parts = v_3d.Name.split()
                                    if len(parts) > 1:
                                        p_3d.Set(parts[1])
                        except Exception as e_vp:
                            erros.append("Aviso VP '%s': %s" % (v_3d.Name, str(e_vp)))

                    if vp_3d and vp_type_id != ElementId.InvalidElementId:
                        try:
                            vp_3d.ChangeTypeId(vp_type_id)
                            doc.Regenerate()
                        except:
                            pass

                    if vp_3d and vp_3d.IsValidObject:
                        try:
                            outline = vp_3d.GetBoxOutline()
                            box_w   = outline.MaximumPoint.X - outline.MinimumPoint.X
                            box_h   = outline.MaximumPoint.Y - outline.MinimumPoint.Y
                            label_h = 0.0
                            label_w = 0.0
                            try:
                                lbl_outline = vp_3d.GetLabelOutline()
                                if lbl_outline is not None:
                                    if lbl_outline.MinimumPoint.Y < outline.MinimumPoint.Y:
                                        label_h = (lbl_outline.MaximumPoint.Y
                                                   - lbl_outline.MinimumPoint.Y)
                                    label_w = (lbl_outline.MaximumPoint.X
                                               - lbl_outline.MinimumPoint.X)
                            except:
                                pass

                            vp_infos.append({
                                "vp":             vp_3d,
                                "vp_id":          vp_3d.Id,
                                "view_id":        v_3d.Id,
                                "box_w":          box_w,
                                "box_h":          box_h,
                                "label_h":        label_h,
                                "label_w":        label_w,
                                "slot_w":         max(box_w, label_w),
                                "slot_h":         box_h + max(label_h, LABEL_H_MIN),
                                "nome":           v_3d.Name,
                                "placeholder_id": placeholder_id,
                                "vp_fantasma_id": vp_fantasma_id,
                                "numero_detalhe": info["numero_detalhe"],
                                "sh_id":          sh.Id,
                            })
                        except Exception as e_dim:
                            erros.append("Aviso dimensao VP '%s': %s" % (v_3d.Name, str(e_dim)))

                t.Commit()
            except Exception as e:
                t.RollBack()
                tg.RollBack()
                forms.alert("Erro ao adicionar viewports: %s" % str(e))
                script.exit()

        if not vp_infos:
            tg.Assimilate()
            forms.alert("Nenhum viewport pôde ser criado.")
            return

        # ---------------------------------------------------------------
        # FASE 3 — SHELF PACKING (cada operacao em transacao propria)
        # ---------------------------------------------------------------

        # ---------------------------------------------------------------
        # FASE 3 — GRID LAYOUT (max 9 por prancha, centralizado)
        # ---------------------------------------------------------------

        def _medir_vp(vp_obj):
            bw = bh = lh = lw = 0.0
            try:
                ol = vp_obj.GetBoxOutline()
                bw = ol.MaximumPoint.X - ol.MinimumPoint.X
                bh = ol.MaximumPoint.Y - ol.MinimumPoint.Y
                try:
                    lbl_ol = vp_obj.GetLabelOutline()
                    if lbl_ol is not None:
                        if lbl_ol.MinimumPoint.Y < ol.MinimumPoint.Y:
                            lh = lbl_ol.MaximumPoint.Y - lbl_ol.MinimumPoint.Y
                        lw = lbl_ol.MaximumPoint.X - lbl_ol.MinimumPoint.X
                except:
                    pass
            except:
                pass
            return bw, bh, lh, lw

        def _recriar_vp_em(sh_dest, vi):
            if vi.get("vp_fantasma_id"):
                try:
                    old_f = doc.GetElement(vi["vp_fantasma_id"])
                    if old_f is not None:
                        doc.Delete(vi["vp_fantasma_id"])
                except:
                    pass
                vi["vp_fantasma_id"] = None

            old_vp = doc.GetElement(vi["vp_id"]) if vi.get("vp_id") else None
            if old_vp is not None and old_vp.IsValidObject:
                try:
                    doc.Delete(vi["vp_id"])
                except:
                    pass
            vi["vp"] = None

            doc.Regenerate()

            ph_id           = vi.get("placeholder_id")
            new_vp          = None
            new_fantasma_id = None

            if ph_id and Viewport.CanAddViewToSheet(doc, sh_dest.Id, ph_id):
                try:
                    vp_f = Viewport.Create(doc, sh_dest.Id, ph_id, XYZ(10, 10, 0))
                    if vp_f:
                        new_fantasma_id = vp_f.Id
                        vista_f = doc.GetElement(ph_id)
                        if vista_f:
                            p_det_f = vista_f.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                            if p_det_f and not p_det_f.IsReadOnly:
                                parts = vista_f.Name.split()
                                if len(parts) > 1:
                                    p_det_f.Set(parts[1])
                    new_vp = Viewport.Create(doc, sh_dest.Id, vi["view_id"], XYZ(0, 0, 0))
                    if new_vp:
                        v3d = doc.GetElement(vi["view_id"])
                        if v3d:
                            p_3d = v3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                            if p_3d and not p_3d.IsReadOnly:
                                parts = v3d.Name.split()
                                if len(parts) > 1:
                                    p_3d.Set(parts[1] + u"\u200b")
                except Exception as e:
                    erros.append("Aviso recriar REF '%s': %s" % (vi["nome"], str(e)))
            else:
                try:
                    new_vp = Viewport.Create(doc, sh_dest.Id, vi["view_id"], XYZ(0, 0, 0))
                    v3d = doc.GetElement(vi["view_id"])
                    if v3d:
                        p_3d = v3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                        if p_3d and not p_3d.IsReadOnly:
                            parts = v3d.Name.split()
                            if len(parts) > 1:
                                p_3d.Set(parts[1])
                except Exception as e:
                    erros.append("Aviso recriar VP '%s': %s" % (vi["nome"], str(e)))

            doc.Regenerate()

            if new_vp and new_vp.IsValidObject and vp_type_id != ElementId.InvalidElementId:
                try:
                    new_vp.ChangeTypeId(vp_type_id)
                    doc.Regenerate()
                except:
                    pass

            if new_vp and new_vp.IsValidObject:
                bw, bh, lh, lw = _medir_vp(new_vp)
                vi["vp"]             = new_vp
                vi["vp_id"]          = new_vp.Id
                vi["box_w"]          = bw
                vi["box_h"]          = bh
                vi["label_h"]        = lh
                vi["label_w"]        = lw
                vi["slot_w"]         = max(bw, lw)
                vi["slot_h"]         = bh + max(lh, LABEL_H_MIN)
                vi["vp_fantasma_id"] = new_fantasma_id
                vi["sh_id"]          = sh_dest.Id
            else:
                erros.append("Aviso: VP nao recriado para '%s'." % vi["nome"])

        # ---------------------------------------------------------------
        # Agrupa os VPs em paginas de ate 9
        # ---------------------------------------------------------------
        #MAX_POR_PRANCHA = 9
        paginas = [vp_infos[i:i + max_por_prancha]
           for i in range(0, len(vp_infos), max_por_prancha)]

        # A primeira prancha ja foi criada; as demais serao criadas no loop
        pranchas = [(sh, ux_min, ux_max, uy_min, uy_max, max_por_prancha)]

        for idx_pag in range(1, len(paginas)):
            with Transaction(doc, "AutoIso: Nova Prancha %d" % (idx_pag + 1)) as t_np:
                t_np.Start()
                try:
                    nova_sh, nx_min, nx_max, ny_min, ny_max, max_pp = nova_prancha()
                    t_np.Commit()
                    pranchas.append((nova_sh, nx_min, nx_max, ny_min, ny_max, max_pp))
                except Exception as e:
                    t_np.RollBack()
                    erros.append("Erro ao criar prancha %d: %s" % (idx_pag + 1, str(e)))
                    pranchas.append(None)

        # ---------------------------------------------------------------
        # Para cada pagina: calcula grid otimo e posiciona
        # ---------------------------------------------------------------
        import math as _math

        for idx_pag, grupo in enumerate(paginas):
            if idx_pag >= len(pranchas) or pranchas[idx_pag] is None:
                continue

            sh_pag, ux_min_p, ux_max_p, uy_min_p, uy_max_p, max_pp = pranchas[idx_pag]
            area_w = ux_max_p - ux_min_p
            area_h = uy_max_p - uy_min_p
            n      = len(grupo)

            # --- Move VPs para a prancha correta (paginas > 0) ----------
            if idx_pag > 0:
                with Transaction(doc, "AutoIso: Mover VPs Prancha %d" % (idx_pag + 1)) as t_mv:
                    t_mv.Start()
                    try:
                        for vi in grupo:
                            _recriar_vp_em(sh_pag, vi)
                        doc.Regenerate()
                        t_mv.Commit()
                    except Exception as e:
                        t_mv.RollBack()
                        erros.append("Erro mover VPs prancha %d: %s" % (idx_pag + 1, str(e)))
                        continue

            # --- Mede o maior slot do grupo (usa o maior para grid uniforme) ---
            slot_w_max = max([vi["slot_w"] for vi in grupo]) if grupo else 0.0
            slot_h_max = max([vi["slot_h"] for vi in grupo]) if grupo else 0.0
            if slot_w_max <= 0 or slot_h_max <= 0:
                erros.append("Aviso: dimensoes invalidas na pagina %d." % (idx_pag + 1))
                continue

           
           
            cols = min(3, n)
            rows = int(_math.ceil(n / float(cols)))

            # --- Espacamento uniforme (distribute evenly) ----------------
            # Espaco sobrando apos alocar os slots
            gap_x = ((area_w - cols * slot_w_max) / (cols + 1)
                     if cols > 0 else 0.0)
            gap_y = ((area_h - rows * slot_h_max) / (rows + 1)
                     if rows > 0 else 0.0)

            # Garante espacamento minimo razoavel (pelo menos 2mm)
            MIN_GAP = 2.0 * mm
            gap_x = max(gap_x, MIN_GAP)
            gap_y = max(gap_y, MIN_GAP)

            # --- Posiciona cada VP no grid (esquerda->direita, cima->baixo) ---
            with Transaction(doc, "AutoIso: Posicionar Grid Prancha %d" % (idx_pag + 1)) as t_pos:
                t_pos.Start()
                try:
                    for i, vi in enumerate(grupo):
                        vp = vi.get("vp")
                        if vp is None and vi.get("vp_id"):
                            vp_elem = doc.GetElement(vi["vp_id"])
                            if vp_elem is not None and vp_elem.IsValidObject:
                                vp = vp_elem
                                vi["vp"] = vp

                        if vp is None or not vp.IsValidObject:
                            erros.append("Aviso: VP invalido '%s', pulando." % vi["nome"])
                            continue

                        col_idx = i % cols
                        row_idx = i // cols

                        # Centro do slot na prancha
                        # Eixo X: gap + (col * (slot_w + gap)) + slot_w/2
                        cx = (ux_min_p
                              + gap_x
                              + col_idx * (slot_w_max + gap_x)
                              + slot_w_max / 2.0)

                        # Eixo Y: topo da area - gap - (row * (slot_h + gap)) - slot_h/2
                        # O topo do box (sem label) e cy + box_h/2
                        cy_slot_top = (uy_max_p
                                       - gap_y
                                       - row_idx * (slot_h_max + gap_y))

                        try:
                            ol = vp.GetBoxOutline()
                            # Queremos que ol.MaximumPoint.Y == cy_slot_top
                            delta_x = cx - (ol.MinimumPoint.X + vi["box_w"] / 2.0)
                            delta_y = cy_slot_top - ol.MaximumPoint.Y
                            DB.ElementTransformUtils.MoveElement(
                                doc, vp.Id, XYZ(delta_x, delta_y, 0))
                            try:
                                vp.LabelOffset = XYZ(0.0, -GAP_LABEL, 0.0)
                            except:
                                pass
                        except Exception as e_mv:
                            erros.append("Erro posicionar '%s': %s" % (vi["nome"], str(e_mv)))

                    doc.Regenerate()
                    t_pos.Commit()
                except Exception as e:
                    t_pos.RollBack()
                    erros.append("Erro ao posicionar grid prancha %d: %s" % (idx_pag + 1, str(e)))

        # Finaliza o TransactionGroup
        tg.Assimilate()

    msg = "Concluido! %d isometrico(s) criado(s) e paginado(s)." % len(vistas_geradas)
    if erros:
        msg += "\n\nAvisos (%d):\n%s" % (len(erros), "\n".join(erros))
    forms.alert(msg)


if __name__ == '__main__':
    executar_fluxo_isometrico()