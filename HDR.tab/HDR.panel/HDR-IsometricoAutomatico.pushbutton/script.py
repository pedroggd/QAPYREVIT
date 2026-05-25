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
    mapping = {
        u'Á': u'A', u'À': u'A', u'Ã': u'A', u'Â': u'A',
        u'É': u'E', u'È': u'E', u'Ê': u'E',
        u'Í': u'I', u'Ì': u'I', u'Î': u'I',
        u'Ó': u'O', u'Ò': u'O', u'Õ': u'O', u'Ô': u'O',
        u'Ú': u'U', u'Ù': u'U', u'Û': u'U',
        u'Ç': u'C',
        u'á': u'a', u'à': u'a', u'ã': u'a', u'â': u'a',
        u'é': u'e', u'è': u'e', u'ê': u'e',
        u'í': u'i', u'ì': u'i', u'î': u'i',
        u'ó': u'o', u'ò': u'o', u'õ': u'o', u'ô': u'o',
        u'ú': u'u', u'ù': u'u', u'û': u'u',
        u'ç': u'c', u'º': u'o', u'ª': u'a',
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
        
        p_pref  = re.escape(prefixo)      if prefixo      else ""
        p_ident = re.escape(identificador) if identificador else ""
        p_suf   = re.escape(sufixo)       if sufixo       else ""
        
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

        self.Title         = "Configuração de Isométricos"
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
        self._lbl(header, "Criar Isométricos+", 16, bold=True, cor=self.COR_AZUL, mg=(0, 0, 0, 2))
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
        btn_ok.Content    = "OK — Iniciar Seleção de Área"
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
        tab.Header = "Configurações"
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
        self._lbl(pnl_num, "Número Inicial:", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_num_inicial = TextBox()
        self.txt_num_inicial.FontSize = 12
        self.txt_num_inicial.Padding  = Thickness(6, 4, 6, 4)
        self.txt_num_inicial.Text     = "01"
        pnl_num.Children.Add(self.txt_num_inicial)
        Grid.SetColumn(pnl_num, 0); g_cont.Children.Add(pnl_num)

        pnl_zero = StackPanel(); pnl_zero.Margin = Thickness(4, 0, 0, 0)
        self._lbl(pnl_zero, "Formato do Número:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_zeros = ComboBox(); self.combo_zeros.FontSize = 11
        for label in ["1 dígito (1, 2, 3)", "2 dígitos (01, 02)", "3 dígitos (001, 002)"]:
            it = ComboBoxItem(); it.Content = label
            self.combo_zeros.Items.Add(it)
        self.combo_zeros.SelectedIndex = 1
        pnl_zero.Children.Add(self.combo_zeros)
        Grid.SetColumn(pnl_zero, 1); g_cont.Children.Add(pnl_zero)

        pnl.Children.Add(g_cont)

        self._lbl(pnl, "Sufixo (ex: '- 1°TIPO' ou vazio):", 11, bold=True, mg=(0, 0, 0, 4))
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

        self._lbl(outer, "Parâmetros do carimbo", 12, bold=True, mg=(0, 0, 0, 4))
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


def executar_fluxo_isometrico():
    titleblock_symbols = get_titleblock_symbols_by_family()
    if not titleblock_symbols:
        forms.alert(
            "Nenhum carimbo encontrado.\nCarregue a família '%s' e tente novamente." % FAMILIA_NOME,
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

    todos_niveis = sorted(
        FilteredElementCollector(doc).OfClass(Level).ToElements(),
        key=lambda n: n.Elevation
    )

    vista_ativa  = doc.ActiveView
    nivel_atual  = vista_ativa.GenLevel

    drafting_vft_id = ElementId.InvalidElementId
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.Drafting:
            drafting_vft_id = vft.Id
            break

    grupos          = []
    contador        = _proximo_contador(prefixo, identificador, num_inicial, sufixo)
    contador_inicial = contador

    while True:
        numero_str    = str(contador).zfill(zeros)
        contador_full = "%s%s" % (identificador, numero_str)
        partes        = [prefixo, contador_full, sufixo]
        nome_preview  = " ".join([p for p in partes if p])

        instrucao = "[%s] Desenhe o retângulo definindo a área do isométrico (ESC para opções)" % nome_preview

        try:
            box = uidoc.Selection.PickBox(PickBoxStyle.Directional, instrucao)
            grupos.append(box)
            contador += 1

        except Exceptions.OperationCanceledException:
            if grupos:
                opcao = forms.alert(
                    "Seleção pausada. %d isométrico(s) na fila.\nO que deseja fazer?" % len(grupos),
                    options=["Finalizar e Criar Pranchas", "Desfazer o ÚLTIMO e Continuar", "Cancelar Script"]
                )
                if opcao == "Desfazer o ÚLTIMO e Continuar":
                    grupos.pop()
                    contador -= 1
                    continue
                elif opcao == "Finalizar e Criar Pranchas":
                    break
                else:
                    script.exit()
            else:
                script.exit()
        except Exception as e:
            forms.alert("Erro na seleção: %s" % str(e))
            break

    if not grupos:
        forms.alert("Nenhum isométrico selecionado. Operação cancelada.")
        script.exit()

    vistas_geradas = []
    erros          = []

    with TransactionGroup(doc, "Gerador de Isométricos") as tg:
        tg.Start()

        for idx, box in enumerate(grupos):
            numero_str    = str(contador_inicial + idx).zfill(zeros)
            contador_full = "%s%s" % (identificador, numero_str)
            partes        = [prefixo, contador_full, sufixo]
            nome_iso      = sanitize_name(" ".join([p for p in partes if p]))

            novo_id        = None
            placeholder_id = None

            with Transaction(doc, "Criar Isométrico: %s" % nome_iso) as t:
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
                        except Exception as e_rastreio:
                            erros.append("Aviso: Falha ao criar referência para '%s': %s"
                                         % (nome_iso, str(e_rastreio)))

                    doc.Regenerate()
                    novo_id = nova.Id
                    t.Commit()

                except Exception as e:
                    erros.append("Erro ao criar '%s': %s" % (nome_iso, str(e)))
                    t.RollBack()

            if novo_id is not None:
                vistas_geradas.append({
                    "id":              novo_id,
                    "placeholder_id":  placeholder_id,
                    "numero_detalhe":  contador_inicial + idx,
                })

        if not vistas_geradas:
            tg.RollBack()
            forms.alert("Nenhuma vista criada.\n" + "\n".join(erros))
            script.exit()

        with Transaction(doc, "Paginação de Pranchas — Isométricos") as t:
            t.Start()
            doc.Regenerate()

            mm                  = 1.0 / 304.8
            margem_esq          = 25.0 * mm
            margem_sup          = 10.0 * mm
            margem_inf          = 15.0 * mm
            MARGEM_CORTE        = 5.0  * mm   # Borda de corte da folha (padrão ABNT)
            espacamento         = 15.0 * mm   # Gap mínimo horizontal entre viewports
            MARGEM_ENTRE_LINHAS = 20.0 * mm   # Gap mínimo vertical entre linhas
            LABEL_H_MIN         = 15.0 * mm   # Altura mínima reservada para o título
            GAP_LABEL           = 5.0  * mm   # Gap entre box e título

            def nova_prancha():
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

                # Tenta detectar ux_max a partir do viewport mais à esquerda da prancha
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

                # MARGEM_CORTE: desconta a borda de corte da folha (padrão ABNT ~5mm)
                # antes de aplicar a margem interna, evitando começar na borda bruta da folha.
                return (
                    sh,
                    bb_sh.Min.X + MARGEM_CORTE + margem_esq,
                    ux_max_calc,
                    bb_sh.Min.Y + MARGEM_CORTE + margem_inf,
                    bb_sh.Max.Y - MARGEM_CORTE - margem_sup,
                )

            sheet, ux_min, ux_max, uy_min, uy_max = nova_prancha()
            vp_infos = []

            for info in vistas_geradas:
                v_3d = doc.GetElement(info["id"])
                if v_3d is None:
                    vp_infos.append(None)
                    continue

                if not Viewport.CanAddViewToSheet(doc, sheet.Id, v_3d.Id):
                    erros.append("Vista '%s' nao pode ser adicionada a prancha." % v_3d.Name)
                    vp_infos.append(None)
                    continue

                placeholder_id = info.get("placeholder_id")
                vp_3d          = None
                vp_fantasma_id = None

                if placeholder_id and Viewport.CanAddViewToSheet(doc, sheet.Id, placeholder_id):
                    vp_fantasma = Viewport.Create(doc, sheet.Id, placeholder_id, XYZ(10, 10, 0))
                    doc.Regenerate()

                    if vp_fantasma:
                        vp_fantasma_id = vp_fantasma.Id
                        vista_fantasma = doc.GetElement(placeholder_id)
                        if vista_fantasma:
                            p_det_f = vista_fantasma.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                            if p_det_f and not p_det_f.IsReadOnly:
                                p_det_f.Set(vista_fantasma.Name.split()[1])

                        vp_3d = Viewport.Create(doc, sheet.Id, v_3d.Id, XYZ(0, 0, 0))
                        doc.Regenerate()
                        if vp_3d:
                            p_3d = v_3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                            if p_3d and not p_3d.IsReadOnly:
                                p_3d.Set(v_3d.Name.split()[1] + u"\u200b")
                else:
                    vp_3d = Viewport.Create(doc, sheet.Id, v_3d.Id, XYZ(0, 0, 0))
                    doc.Regenerate()
                    p_3d = v_3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                    if p_3d and not p_3d.IsReadOnly:
                        p_3d.Set(v_3d.Name.split()[1])

                if vp_3d and vp_type_id != ElementId.InvalidElementId:
                    vp_3d.ChangeTypeId(vp_type_id)
                    doc.Regenerate()

                if vp_3d:
                    outline = vp_3d.GetBoxOutline()
                    box_w   = outline.MaximumPoint.X - outline.MinimumPoint.X
                    box_h   = outline.MaximumPoint.Y - outline.MinimumPoint.Y
                    label_h = 0.0
                    label_w = 0.0
                    try:
                        lbl_outline = vp_3d.GetLabelOutline()
                        if lbl_outline is not None:
                            if lbl_outline.MinimumPoint.Y < outline.MinimumPoint.Y:
                                label_h = lbl_outline.MaximumPoint.Y - lbl_outline.MinimumPoint.Y
                            # Título pode ser mais largo que o box: slot_w precisa considerar isso
                            label_w = lbl_outline.MaximumPoint.X - lbl_outline.MinimumPoint.X
                    except:
                        pass

                    vp_infos.append({
                        "vp":             vp_3d,
                        "view_id":        v_3d.Id,
                        "box_w":          box_w,
                        "box_h":          box_h,
                        "label_h":        label_h,
                        "label_w":        label_w,
                        "slot_w":         max(box_w, label_w),        # ← corrigido
                        "slot_h":         box_h + max(label_h, LABEL_H_MIN),  # ← corrigido
                        "nome":           v_3d.Name,
                        "placeholder_id": placeholder_id,
                        "vp_fantasma_id": vp_fantasma_id,
                        "numero_detalhe": info["numero_detalhe"],
                    })
                else:
                    vp_infos.append(None)

            vp_infos = [vi for vi in vp_infos if vi is not None]

            if not vp_infos:
                t.Commit()
                tg.Assimilate()
                forms.alert("Nenhum viewport pôde ser criado.")
                return

            import math as _math

            area_w = ux_max - ux_min
            area_h = uy_max - uy_min
            ref_sw = max(vi["slot_w"] for vi in vp_infos)
            ref_sh = max(vi["slot_h"] for vi in vp_infos)

            ncols = max(1, int((area_w + espacamento) / (ref_sw + espacamento)))
            nrows = max(1, int((area_h + MARGEM_ENTRE_LINHAS) / (ref_sh + MARGEM_ENTRE_LINHAS)))
            cap   = ncols * nrows
            total = len(vp_infos)

            n_pranchas = max(1, int(_math.ceil(total / float(cap))))

            grupos_prancha = []
            restante = list(vp_infos)
            for _p in range(n_pranchas):
                qtd = int(_math.ceil(len(restante) / float(n_pranchas - _p)))
                grupos_prancha.append(restante[:qtd])
                restante = restante[qtd:]

            def _recriar_vp_na_prancha(sh_dest, vi):
                if vi.get("vp_fantasma_id"):
                    try:
                        doc.Delete(vi["vp_fantasma_id"])
                        doc.Regenerate()
                    except:
                        pass
                try:
                    doc.Delete(vi["vp"].Id)
                    doc.Regenerate()
                except:
                    pass

                ph_id           = vi.get("placeholder_id")
                new_vp          = None
                new_fantasma_id = None

                if ph_id and Viewport.CanAddViewToSheet(doc, sh_dest.Id, ph_id):
                    vp_f = Viewport.Create(doc, sh_dest.Id, ph_id, XYZ(10, 10, 0))
                    doc.Regenerate()
                    if vp_f:
                        new_fantasma_id = vp_f.Id
                        vista_fantasma  = doc.GetElement(ph_id)
                        if vista_fantasma:
                            p_det_f = vista_fantasma.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                            if p_det_f and not p_det_f.IsReadOnly:
                                p_det_f.Set(vista_fantasma.Name.split()[1])

                        new_vp = Viewport.Create(doc, sh_dest.Id, vi["view_id"], XYZ(0, 0, 0))
                        doc.Regenerate()
                        if new_vp:
                            vista_3d = doc.GetElement(vi["view_id"])
                            if vista_3d:
                                p_3d = vista_3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                                if p_3d and not p_3d.IsReadOnly:
                                    p_3d.Set(vista_3d.Name.split()[1] + u"\u200b")
                else:
                    new_vp = Viewport.Create(doc, sh_dest.Id, vi["view_id"], XYZ(0, 0, 0))
                    doc.Regenerate()
                    vista_3d = doc.GetElement(vi["view_id"])
                    if vista_3d:
                        p_3d = vista_3d.get_Parameter(BuiltInParameter.VIEWPORT_DETAIL_NUMBER)
                        if p_3d and not p_3d.IsReadOnly:
                            p_3d.Set(vista_3d.Name.split()[1])

                if new_vp and vp_type_id != ElementId.InvalidElementId:
                    new_vp.ChangeTypeId(vp_type_id)
                    doc.Regenerate()

                if new_vp:
                    ol = new_vp.GetBoxOutline()
                    bw = ol.MaximumPoint.X - ol.MinimumPoint.X
                    bh = ol.MaximumPoint.Y - ol.MinimumPoint.Y
                    lh = 0.0
                    lw = 0.0
                    try:
                        lbl_ol = new_vp.GetLabelOutline()
                        if lbl_ol is not None:
                            if lbl_ol.MinimumPoint.Y < ol.MinimumPoint.Y:
                                lh = lbl_ol.MaximumPoint.Y - lbl_ol.MinimumPoint.Y
                            lw = lbl_ol.MaximumPoint.X - lbl_ol.MinimumPoint.X
                    except:
                        pass
                    vi["vp"]             = new_vp
                    vi["box_w"]          = bw
                    vi["box_h"]          = bh
                    vi["label_h"]        = lh
                    vi["label_w"]        = lw
                    vi["slot_w"]         = max(bw, lw)               # ← corrigido
                    vi["slot_h"]         = bh + max(lh, LABEL_H_MIN) # ← corrigido
                    vi["vp_fantasma_id"] = new_fantasma_id

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

                linhas_g = [grupo[i:i + ncols] for i in range(0, len(grupo), ncols)]
                n_lin    = len(linhas_g)

                alt_lins = [max(vi["slot_h"] for vi in ln) for ln in linhas_g]
                tot_h    = sum(alt_lins)

                GAP_Y_MAX = 40.0 * mm
                if n_lin > 1:
                    gap_y = min(GAP_Y_MAX, max(MARGEM_ENTRE_LINHAS,
                                (area_h_p - tot_h) / float(n_lin - 1)))
                else:
                    gap_y = 0.0

                cur_y_p = uy_max_p

                for l_idx, linha in enumerate(linhas_g):
                    rh_max = alt_lins[l_idx]
                    n_col  = len(linha)
                    tot_w  = sum(vi["slot_w"] for vi in linha)

                    GAP_X_MAX = 20.0 * mm
                    if n_col > 1:
                        gap_x = min(GAP_X_MAX, max(espacamento,
                                    (area_w_p - tot_w) / float(n_col - 1)))
                    else:
                        gap_x = 0.0

                    bloco_w = tot_w + gap_x * (n_col - 1)
                    cur_x_p = ux_min_p + (area_w_p - bloco_w) / 2.0
                    base_y  = cur_y_p - rh_max

                    for vi in linha:
                        try:
                            vp = vi["vp"]
                            if vp is None:
                                continue
                            ol = vp.GetBoxOutline()
                            DB.ElementTransformUtils.MoveElement(
                                doc, vp.Id,
                                XYZ(cur_x_p - ol.MinimumPoint.X,
                                    base_y  - ol.MinimumPoint.Y, 0))
                            # Título sempre logo abaixo do box, com gap fixo —
                            # igual ao script de detalhes 2D (LabelOffset relativo ao centro do box)
                            vp.LabelOffset = XYZ(0.0, -GAP_LABEL, 0.0)
                            doc.Regenerate()
                        except Exception as e:
                            erros.append("Erro ao posicionar '%s': %s" % (vi["nome"], str(e)))
                        cur_x_p += vi["slot_w"] + gap_x

                    cur_y_p -= (rh_max + gap_y)

            t.Commit()

        tg.Assimilate()

    msg = "Concluído! %d isométrico(s) criado(s) e paginado(s)." % len(vistas_geradas)
    if erros:
        msg += "\n\nAvisos (%d):\n%s" % (len(erros), "\n".join(erros))
    forms.alert(msg)


if __name__ == '__main__':
    executar_fluxo_isometrico()