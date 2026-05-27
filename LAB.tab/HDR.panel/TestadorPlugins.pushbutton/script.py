# -*- coding: utf-8 -*-

__title__ = "TESTE.APENAS+"
__author__ = "PyRevit Plugin"

import clr
import System
import math
import re

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("Microsoft.Office.Interop.Excel")

from System.Windows import Window, Thickness, GridLength, GridUnitType
from System.Windows.Controls import (
    StackPanel, Label, TextBox, Button, ScrollViewer,
    Separator, Grid, ColumnDefinition, RowDefinition, ComboBox, ComboBoxItem,
    TabControl, TabItem, CheckBox, WrapPanel
)
from System.Windows.Media import SolidColorBrush, Color
from System.Windows import FontWeights, ResizeMode, WindowStartupLocation
from System.Windows.Controls import ScrollBarVisibility

import Microsoft.Office.Interop.Excel as Excel
from pyrevit import forms, revit, script
from Autodesk.Revit.DB import (
    FilteredElementCollector, FamilySymbol, ViewFamilyType, ViewFamily,
    ViewSheet, Transaction, BuiltInCategory, BuiltInParameter,
    ViewDuplicateOption, Viewport, XYZ,
    ViewType, StorageType, BoundingBoxXYZ, ElementId, Options, Line, GeometryInstance
)
import Autodesk.Revit.DB as DB
import Autodesk.Revit.Exceptions as Exceptions

logger = script.get_logger()
doc    = revit.doc

ORDEM_TITLEBLOCK = [
    "A4", "A3 - Retrato", "A3 - Paisagem", "A2 - Retrato", "A2 - Paisagem",
    "A1 - Retrato", "A1 - Paisagem", "A0 - Retrato", "A0 - Paisagem",
]

MARGEM_ESQ_MM   = 25.0
MARGEM_DIR_MM   = 180.0
MARGEM_SUP_MM   = 10.0
MARGEM_INF_MM   = 10.0

ESCALAS_VALIDAS = [50, 75, 100]

SIGLA_TO_PARAM = {
    "EE":   "Projeto de Elétrica",
    "PE":   "Projeto de PE",
    "EP":   "Projeto de PE",
    "TE":   "Projeto de Comunicação",
    "TP":   "Projeto de TP",
    "SPDA": "Projeto de SPDA",
    "SDAI": "Projeto de SDAI"
}

PARAM_DISC_MATCH = {
    "ARQ": ["ARQUITETURA", "ARQ"],
    "EST": ["ESTRUTURA", "EST"],
    "ELE": ["ELETRICA", "ELE"],
    "HID": ["HIDRAULICA", "ESGOTO", "HID"],
    "APL": ["PLUVIAIS", "AGUAS PLUVIAIS", "APL"],
    "CLI": ["CLIMATIZACAO", "CLI"],
    "PCI": ["INCENDIO", "PREVENCAO", "PCI"],
    "GAS": ["GAS"],
    "TEL": ["TELECOMUNICACOES", "TEL"],
    "IMP": ["IMPERMEABILIZACAO", "IMP"],
    "PAI": ["PAISAGISMO", "PAI"],
    "INT": ["INTERIORES", "INT"],

    "HI":  ["HIDRAULICA", "PH", "HI"],
    "PH":  ["HIDRAULICA", "PH", "HI"],
    "INC": ["INCENDIO", "PI", "INC"],
    "PI":  ["INCENDIO", "PI", "INC"],
    "AC":  ["AVAC", "MECANICA", "PM", "AC"],
    "PM":  ["AVAC", "MECANICA", "PM", "AC"],
    "PT":  ["COMUNICACAO", "TELECOM", "TE", "PT"]
}

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

NOMES_PARAM_TITULO = ["HDR-Título", "HDR-Titulo", "TÍTULO", "TITULO"]

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

def set_param_titulo(sheet, tb_instance, descricao):
    for nome in NOMES_PARAM_TITULO:
        p_sheet = sheet.LookupParameter(nome)
        if p_sheet and not p_sheet.IsReadOnly:
            p_sheet.Set(descricao)
            return True
        if tb_instance:
            p_tb = tb_instance.LookupParameter(nome)
            if p_tb and not p_tb.IsReadOnly:
                p_tb.Set(descricao)
                return True
    return False


def set_viewport_titulo(viewport, view_name):
    nomes_possiveis = ["HDR Título de Vista", "HDR - Título de Vista", "Título de Vista", "Title", "View Title"]
    for nome_param in nomes_possiveis:
        try:
            p_vp = viewport.LookupParameter(nome_param)
            if p_vp and not p_vp.IsReadOnly:
                if p_vp.StorageType == StorageType.String:
                    p_vp.Set(view_name)
                    return True
        except:
            continue
    return False

def importar_lista_mestra_banco():
    path = forms.pick_file(file_ext="xlsx", title="Selecione a Lista Mestra Excel")
    if not path: return None

    # Mapeamento fixo
    CONFIG_ABAS = {
        "C.I CIVIL": {"COD": 7, "PROJ": 8, "DISC": 9, "ETAP": 10, "PRANCHA": 11, "NOME": 12},
        "C.I ELE":   {"COD": 11, "PROJ": 7, "DISC": 8, "ETAP": 12, "PRANCHA": 13, "NOME": 16},
        "C.I MEC":   {"COD": 11, "PROJ": 7, "DISC": 8, "ETAP": 12, "PRANCHA": 13, "NOME": 16}
    }

    aba_selecionada = forms.alert("Selecione a aba:", options=CONFIG_ABAS.keys())
    if not aba_selecionada: return {}, []

    col = CONFIG_ABAS[aba_selecionada]
    app_excel = Excel.ApplicationClass()
    app_excel.Visible = False
    
    banco_pranchas = {}
    siglas_encontradas = set()
    
    # Coleta pranchas já existentes para evitar duplicidade
    todas_pranchas = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    numeros_existentes = set([p.SheetNumber for p in todas_pranchas])

    def limpar(val):
        if val is None: return ""
        if isinstance(val, float) and val.is_integer(): return str(int(val))
        return str(val).strip()

    try:
        wb = app_excel.Workbooks.Open(path, False, True)
        ws = wb.Sheets(aba_selecionada)
        
        # VARREDURA TOTAL: Vai da linha 4 até a 1000, sem parar no primeiro erro
        for linha in range(3, 1000):
            codigo      = limpar(ws.Cells[linha, col["COD"]].Value2)
            projeto     = limpar(ws.Cells[linha, col["PROJ"]].Value2)
            disciplina  = limpar(ws.Cells[linha, col["DISC"]].Value2)
            etapa       = limpar(ws.Cells[linha, col["ETAP"]].Value2)
            prancha_raw = limpar(ws.Cells[linha, col["PRANCHA"]].Value2)
            nome_prancha= limpar(ws.Cells[linha, col["NOME"]].Value2)

            # Só ignora se não tiver NADA na linha
            if not codigo and not prancha_raw and not nome_prancha:
                continue
            
            # Se tiver prancha ou código, processa (independente da disciplina)
            if prancha_raw:
                prancha = prancha_raw.zfill(2) if prancha_raw.isdigit() else prancha_raw
                titulo_estruturado = "{}-{}-{}-{}-{}".format(codigo, projeto, disciplina, etapa, prancha)
                
                num_final = prancha
                # Garante exclusividade no Revit
                while num_final in numeros_existentes: num_final += "*"
                numeros_existentes.add(num_final)
                
                sigla_limpa = disciplina.upper()
                if sigla_limpa: siglas_encontradas.add(sigla_limpa)

                banco_pranchas[num_final] = {
                    "numero": num_final,
                    "descricao": nome_prancha,      
                    "nome_arquivo": titulo_estruturado, 
                    "sigla": sigla_limpa
                }
                
    except Exception as e:
        forms.alert("Erro fatal na leitura: {}".format(e))
    finally:
        wb.Close(False)
        app_excel.Quit()

    return banco_pranchas, list(siglas_encontradas)

def get_all_views():
    excluded = {ViewType.DrawingSheet, ViewType.ProjectBrowser,
                ViewType.SystemBrowser, ViewType.Undefined}
    views = {}
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate:
            continue
        if v.ViewType in excluded:
            continue
        if not v.CanViewBeDuplicated(ViewDuplicateOption.WithDetailing):
            continue
        label = "[{}] {}".format(v.ViewType.ToString(), v.Name)
        views[label] = v
    return views

def get_all_view_templates():
    templates = {"(Nenhum)": None}
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.IsTemplate:
            templates[v.Name] = v
    return templates

def get_all_legends():
    legends = {}
    for v in FilteredElementCollector(doc).OfClass(DB.View):
        if v.ViewType == ViewType.Legend and not v.IsTemplate:
            legends[v.Name] = v
    return legends

def get_all_viewport_types():
    tipos = {}
    try:
        instancias = FilteredElementCollector(doc).OfClass(Viewport).ToElements()
        for inst in instancias:
            try:
                tipo = doc.GetElement(inst.GetTypeId())
                if tipo is None: continue
                p_nome = tipo.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                nome = p_nome.AsString() if p_nome else (tipo.Name if tipo.Name else "?")
                if nome not in tipos: tipos[nome] = tipo.Id
            except: pass
    except: pass

    if not tipos:
        try:
            cat_id = DB.Category.GetCategory(doc, BuiltInCategory.OST_Viewports)
            for vpt in FilteredElementCollector(doc).OfClass(DB.ElementType):
                try:
                    if vpt.Category is not None and vpt.Category.Id == cat_id.Id:
                        p_nome = vpt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                        nome = p_nome.AsString() if p_nome else (vpt.Name if vpt.Name else "?")
                        if nome not in tipos: tipos[nome] = vpt.Id
                except: pass
        except: pass

    if not tipos:
        for vpt in FilteredElementCollector(doc).OfClass(DB.ElementType):
            try:
                familia = (vpt.FamilyName or "").lower()
                if "viewport" in familia or "visor" in familia:
                    p_nome = vpt.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                    nome = p_nome.AsString() if p_nome else (vpt.Name if vpt.Name else "?")
                    if nome not in tipos: tipos[nome] = vpt.Id
            except: pass

    if not tipos: tipos["(Padrao - sem alteracao)"] = None
    return tipos

def get_titleblocks_dict():
    tbs = {}
    for symbol in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).OfClass(FamilySymbol):
        p_nome = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        nome_tipo = p_nome.AsString() if p_nome else symbol.Name
        label = "{} : {}".format(symbol.Family.Name, nome_tipo)
        tbs[label] = symbol
    return tbs

def get_todos_titleblocks_com_area_util():
    mm_to_ft = 1.0 / 304.8
    DIMS_MM = {
        "A0 - Paisagem":  (1189.0, 841.0),
        "A0 - Retrato":   (841.0, 1189.0),
        "A0":             (1189.0, 841.0),
        "A1 - Paisagem":  (841.0, 594.0),
        "A1 - Retrato":   (594.0, 841.0),
        "A1":             (841.0, 594.0),
        "A2 - Paisagem":  (594.0, 420.0),
        "A2 - Retrato":   (420.0, 594.0),
        "A2":             (594.0, 420.0),
        "A3 - Paisagem":  (420.0, 297.0),
        "A3 - Retrato":   (297.0, 420.0),
        "A3":             (420.0, 297.0),
        "A4 - Paisagem":  (297.0, 210.0),
        "A4 - Retrato":   (210.0, 297.0),
        "A4":             (210.0, 297.0),
    }

    marg_esq = MARGEM_ESQ_MM * mm_to_ft
    marg_dir = MARGEM_DIR_MM * mm_to_ft
    marg_sup = MARGEM_SUP_MM * mm_to_ft
    marg_inf = MARGEM_INF_MM * mm_to_ft

    resultados = []

    for symbol in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).OfClass(FamilySymbol):
        try:
            p_nome = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            nome_tipo = p_nome.AsString() if p_nome else symbol.Name

            total_w_ft = None
            total_h_ft = None

            dims = DIMS_MM.get(nome_tipo)
            if dims is None:
                nome_upper = nome_tipo.upper()
                for k, v in DIMS_MM.items():
                    if k.upper() in nome_upper or nome_upper in k.upper():
                        dims = v
                        break
            if dims:
                total_w_ft = dims[0] * mm_to_ft
                total_h_ft = dims[1] * mm_to_ft

            if total_w_ft is None or total_h_ft is None:
                nomes_w = ["Largura", "Width", "W", "Sheet Width", "Folha Largura"]
                nomes_h = ["Altura",  "Height", "H", "Sheet Height", "Folha Altura"]
                for nw in nomes_w:
                    p = symbol.LookupParameter(nw)
                    if p and p.StorageType == StorageType.Double and p.AsDouble() > 0:
                        total_w_ft = p.AsDouble()
                        break
                for nh in nomes_h:
                    p = symbol.LookupParameter(nh)
                    if p and p.StorageType == StorageType.Double and p.AsDouble() > 0:
                        total_h_ft = p.AsDouble()
                        break

            if total_w_ft is None or total_h_ft is None:
                p_w = symbol.get_Parameter(BuiltInParameter.SHEET_WIDTH)
                p_h = symbol.get_Parameter(BuiltInParameter.SHEET_HEIGHT)
                if p_w and p_w.AsDouble() > 0: total_w_ft = p_w.AsDouble()
                if p_h and p_h.AsDouble() > 0: total_h_ft = p_h.AsDouble()

            if not total_w_ft or not total_h_ft: continue

            util_w = total_w_ft - marg_esq - marg_dir
            util_h = total_h_ft - marg_sup - marg_inf

            if util_w <= 0 or util_h <= 0: continue

            resultados.append((symbol, util_w, util_h, nome_tipo))
        except:
            continue

    def _sort_key(item):
        nome = item[3]
        try: idx = ORDEM_TITLEBLOCK.index(nome)
        except ValueError: idx = len(ORDEM_TITLEBLOCK)
        area = item[1] * item[2]
        return (idx, area)

    resultados.sort(key=_sort_key)
    return resultados


def escolher_melhor_titleblock(modelo_w, modelo_h, tb_preferido, todos_tbs, nome_vista="", w_leg_max=0, h_leg_total=0):
    vista_e_paisagem = modelo_w >= modelo_h

    def sort_key_orientado(item):
        tb_sym, util_w, util_h, nome_tipo = item
        formato_e_paisagem = util_w >= util_h
        orientacao_errada = (vista_e_paisagem != formato_e_paisagem)
        area = util_w * util_h
        try:
            idx_base = ORDEM_TITLEBLOCK.index(nome_tipo)
        except ValueError:
            idx_base = len(ORDEM_TITLEBLOCK)
        return (orientacao_errada, idx_base, area)

    todos_tbs_ordenados = sorted(todos_tbs, key=sort_key_orientado)

    melhor_tb     = None
    melhor_escala = None

    for tb_sym, util_w, util_h, nome_tipo in todos_tbs_ordenados:
        for escala in sorted(ESCALAS_VALIDAS):
            vista_w = modelo_w / float(escala)
            vista_h = modelo_h / float(escala)
            
            cabe_normal   = (vista_w <= util_w and vista_h <= util_h)
            cabe_rotacion = (vista_h <= util_w and vista_w <= util_h)
            
            if (cabe_normal or cabe_rotacion) and h_leg_total <= util_h:
                melhor_tb     = tb_sym
                melhor_escala = escala
                break
        if melhor_tb is not None:
            break

    if melhor_tb is None:
        melhor_tb     = tb_preferido if tb_preferido else todos_tbs_ordenados[-1][0]
        melhor_escala = sorted(ESCALAS_VALIDAS)[-1]
        tb_info = next((t for t in todos_tbs if t[0].Id == melhor_tb.Id), None)
        nome_tb = tb_info[3] if tb_info else "?"
        aviso = (
            u"AVISO: Vista '{}' e legendas não couberam em nenhum formato disponivel. "
            u"Usando '{}' com escala 1:{}.".format(nome_vista, nome_tb, melhor_escala)
        )
        return melhor_tb, melhor_escala, aviso

    aviso = None
    if tb_preferido and melhor_tb.Id != tb_preferido.Id:
        tb_info_preferido = next((t for t in todos_tbs if t[0].Id == tb_preferido.Id), None)
        tb_info_escolhido = next((t for t in todos_tbs if t[0].Id == melhor_tb.Id), None)
        nome_escolhido    = tb_info_escolhido[3] if tb_info_escolhido else "?"
        nome_preferido    = tb_info_preferido[3] if tb_info_preferido else "?"
        aviso = (
            u"INFO: Vista '{}' alocada em '{}' (escala 1:{}) "
            u"em vez de '{}' selecionado para acomodar os elementos.".format(
                nome_vista, nome_escolhido, melhor_escala, nome_preferido
            )
        )
    elif not tb_preferido:
        tb_info_escolhido = next((t for t in todos_tbs if t[0].Id == melhor_tb.Id), None)
        nome_escolhido    = tb_info_escolhido[3] if tb_info_escolhido else "?"
        aviso = (
            u"INFO: Vista '{}' auto-alocada em '{}' (escala 1:{}).".format(
                nome_vista, nome_escolhido, melhor_escala
            )
        )

    return melhor_tb, melhor_escala, aviso


class CriarPranchasWindow(Window):
    COR_AZUL    = Color.FromRgb(30, 80, 160)
    COR_VERDE   = Color.FromRgb(22, 160, 80)
    COR_LARANJA = Color.FromRgb(180, 90, 0)
    COR_CINZA   = Color.FromRgb(120, 120, 120)
    COR_VERMELHO= Color.FromRgb(220, 50, 50)
    COR_FUNDO   = Color.FromRgb(245, 245, 245)
    COR_BRANCO  = Color.FromRgb(255, 255, 255)

    def __init__(self, views_dict, templates_dict, legends_dict, viewport_types_dict, titleblocks_dict, vistas_geradas=None):
        self.views_dict           = views_dict
        self.templates_dict       = templates_dict
        self.legends_dict         = legends_dict
        self.viewport_types_dict  = viewport_types_dict
        self.titleblocks_dict     = titleblocks_dict
        self.vistas_geradas       = vistas_geradas or []
        
        self.banco_excel           = {} 
        self.pranchas_selecionadas = set() 
        self.campos_pranchas       = []
        self.linhas_config         = []
        self.resultado             = None
        self._mapeamento_ids       = {}
        self.legendas_por_sigla    = {} 
        self.legendas_manuais_temp = []
        self._vp_type_ids_list     = []

        self.Title = "Criar Pranchas+ Definitivo - PLAENGE"
        self.Width = 1050
        self.MinHeight = 700
        self.MaxHeight = 950
        self.ResizeMode = ResizeMode.NoResize
        self.WindowStartupLocation = WindowStartupLocation.CenterScreen
        self.Background = SolidColorBrush(self.COR_FUNDO)

        root = Grid()
        root.Margin = Thickness(16, 14, 16, 14)

        root.RowDefinitions.Add(self._row(GridLength.Auto))
        root.RowDefinitions.Add(self._row(GridLength(1, GridUnitType.Star)))
        root.RowDefinitions.Add(self._row(GridLength.Auto))

        header_panel = StackPanel()
        self._lbl(header_panel, "Criar Pranchas+", 17, bold=True, cor=self.COR_AZUL, mg=(0,0,0,2))
        self._lbl(header_panel, "Fluxo: Busca e Seleção no Excel ou Criação Manual", 10, cor=self.COR_CINZA, mg=(0,0,0,10))
        header_panel.Children.Add(self._sep(10))
        Grid.SetRow(header_panel, 0)
        root.Children.Add(header_panel)

        self.tabs = TabControl()
        self.tabs.FontSize = 12
        self.tabs.Margin = Thickness(0, 0, 0, 12)
        Grid.SetRow(self.tabs, 1)
        root.Children.Add(self.tabs)

        self._build_tab_config()
        self._build_tab_manual()
        self._build_tab_pranchas()
        self._build_tab_carimbo()

        btn = self._btn("Criar Pranchas Selecionadas", self.COR_VERDE, size=13, bold=True)
        btn.Padding = Thickness(0, 10, 0, 10)
        btn.Click  += self.on_criar
        Grid.SetRow(btn, 2)
        root.Children.Add(btn)

        self.Content = root

    def _row(self, height):
        r = RowDefinition()
        r.Height = height
        return r

    def _col(self, width):
        c = ColumnDefinition()
        c.Width = width
        return c

    def _build_tab_config(self):
        tab = TabItem()
        tab.Header = "Configurações / Importar Excel"
        outer = StackPanel()
        outer.Margin = Thickness(12, 12, 12, 12)

        g_top = Grid()
        g_top.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g_top.ColumnDefinitions.Add(self._col(GridLength.Auto))
        
        info_panel = StackPanel()
        self._lbl(info_panel, "1. Importe a Lista Mestra", 12, bold=True, mg=(0,0,0,4))
        self._lbl(info_panel, "O plugin vai ler as siglas (EE, PE, TE...) e pedir para você mapear os View Templates.", 10, cor=self.COR_CINZA, mg=(0,0,0,0))
        Grid.SetColumn(info_panel, 0)
        g_top.Children.Add(info_panel)

        btn_excel = self._btn("Importar Lista Mestra Excel", self.COR_LARANJA, size=11, bold=True)
        btn_excel.Click += self.on_importar_excel
        Grid.SetColumn(btn_excel, 1)
        g_top.Children.Add(btn_excel)
        
        outer.Children.Add(g_top)
        outer.Children.Add(self._sep(10))

        self._lbl(outer, "Sigla do Excel  ->  View Template  &  Legendas", 12, bold=True, mg=(0,0,0,6))

        scroll = ScrollViewer()
        scroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll.MaxHeight = 400
        scroll.Margin = Thickness(0, 0, 0, 10)
        self.painel_config = StackPanel()
        scroll.Content = self.painel_config
        outer.Children.Add(scroll)

        self._lbl(self.painel_config, "Importe o Excel para ver as siglas...", 11, cor=self.COR_CINZA, mg=(10,10,0,0))

        tab.Content = outer
        self.tabs.Items.Add(tab)
        
    def _build_tab_manual(self):
        tab = TabItem()
        tab.Header = "Criação Manual"
        
        pnl = StackPanel()
        pnl.Margin = Thickness(12)
        
        self._lbl(pnl, "Criar Prancha e Vista Manualmente", 12, bold=True, mg=(0,0,0,10))
        self._lbl(pnl, "Use esta aba se o Excel estiver indisponível ou se precisar criar pranchas extras.", 10, cor=self.COR_CINZA, mg=(0,0,0,10))
        
        g = Grid()
        g.ColumnDefinitions.Add(self._col(GridLength(130)))
        g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        
        row_idx = 0
        def add_row(label, element):
            g.RowDefinitions.Add(self._row(GridLength.Auto))
            l = self._lbl(None, label, 11, mg=(0,8,8,8))
            Grid.SetRow(l, row_idx)
            Grid.SetColumn(l, 0)
            g.Children.Add(l)
            
            element.Margin = Thickness(0,0,0,8)
            Grid.SetRow(element, row_idx)
            Grid.SetColumn(element, 1)
            g.Children.Add(element)
            return row_idx + 1

        self.txt_man_num = TextBox()
        self.txt_man_num.Padding = Thickness(4)
        row_idx = add_row("Número:", self.txt_man_num)
        
        self.txt_man_nome = TextBox()
        self.txt_man_nome.Padding = Thickness(4)
        row_idx = add_row("Nome da Prancha:", self.txt_man_nome)
        
        self.txt_man_sigla = TextBox()
        self.txt_man_sigla.Padding = Thickness(4)
        row_idx = add_row("Disciplina (Sigla):", self.txt_man_sigla)

        self.cb_man_tb = ComboBox()
        for tb_name in sorted(self.titleblocks_dict.keys()):
            i = ComboBoxItem()
            i.Content = tb_name
            self.cb_man_tb.Items.Add(i)
        if self.cb_man_tb.Items.Count > 0:
            self.cb_man_tb.SelectedIndex = 0
        row_idx = add_row("Formato da Prancha:", self.cb_man_tb)
        
        self.cb_man_base = ComboBox()
        for v in sorted(self.views_dict.keys()):
            i = ComboBoxItem()
            i.Content = v
            self.cb_man_base.Items.Add(i)
        if self.cb_man_base.Items.Count > 0:
            self.cb_man_base.SelectedIndex = 0
        row_idx = add_row("Vista Base (Origem):", self.cb_man_base)
        
        self.cb_man_temp = ComboBox()
        self.mapa_man_temp = []
        for t in sorted(self.templates_dict.keys()):
            i = ComboBoxItem()
            i.Content = t
            self.cb_man_temp.Items.Add(i)
            self.mapa_man_temp.append(self.templates_dict[t].Id if self.templates_dict[t] else None)
        if self.cb_man_temp.Items.Count > 0:
            self.cb_man_temp.SelectedIndex = 0
        row_idx = add_row("View Template:", self.cb_man_temp)

        self.btn_man_legs = self._btn("Legendas (0)", self.COR_CINZA, size=11)
        self.btn_man_legs.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
        self.btn_man_legs.Click += self.on_selecionar_legendas_manual
        row_idx = add_row("Legendas:", self.btn_man_legs)
        
        pnl.Children.Add(g)
        
        btn_add = self._btn("Adicionar à Fila de Criação (Aba Pranchas)", self.COR_AZUL, size=11, bold=True)
        btn_add.Margin = Thickness(0, 15, 0, 0)
        btn_add.Click += self.on_add_manual
        pnl.Children.Add(btn_add)
        
        tab.Content = pnl
        self.tabs.Items.Add(tab)

    def on_selecionar_legendas_manual(self, sender, args):
        opcoes = sorted(self.legends_dict.keys())
        sel = forms.SelectFromList.show(opcoes, title="Legendas para Vista Manual", multiselect=True)
        if sel is not None:
            self.legendas_manuais_temp = sel
            sender.Content = "Legendas ({})".format(len(sel))
        
    def on_add_manual(self, sender, args):
        num = self.txt_man_num.Text.strip()
        nome = self.txt_man_nome.Text.strip()
        sigla = self.txt_man_sigla.Text.strip()
        
        if not num:
            forms.alert("O Número da prancha é obrigatório.")
            return
            
        idx_temp = self.cb_man_temp.SelectedIndex
        temp_id = self.mapa_man_temp[idx_temp] if idx_temp >= 0 else None
        
        base_view_str = self.cb_man_base.SelectedItem.Content if self.cb_man_base.SelectedItem else None
        tb_str = self.cb_man_tb.SelectedItem.Content if self.cb_man_tb.SelectedItem else None
        
        self._add_linha_prancha_criacao(
            numero=num,
            nome_arquivo=nome,
            sigla=sigla,
            manual_template_id=temp_id,
            is_manual=True,
            manual_base_view_str=base_view_str,
            manual_tb_str=tb_str,
            legendas_iniciais=self.legendas_manuais_temp
        )
        
        forms.alert("Prancha adicionada com sucesso na aba 'Pranchas'.")
        self.txt_man_num.Text = ""
        self.txt_man_nome.Text = ""
        self.legendas_manuais_temp = []
        self.btn_man_legs.Content = "Legendas (0)"

    def on_importar_excel(self, sender, args):
        retorno = importar_lista_mestra_banco()
        if not retorno or not retorno[0]: return
            
        self.banco_excel, siglas = retorno
        self.pranchas_selecionadas.clear()
        
        self.painel_config.Children.Clear()
        self.linhas_config = []
        self._mapeamento_ids = {}
        
        if not siglas:
            self._lbl(self.painel_config, "Nenhuma sigla de disciplina encontrada na coluna C do Excel.", 11, cor=Color.FromRgb(200,50,50))
            return
            
        siglas.sort()
        for sigla in siglas: self._criar_linha_config(sigla)
            
        if hasattr(self, 'combo_filtro'):
            self.combo_filtro.Items.Clear()
            item_todas = ComboBoxItem()
            item_todas.Content = "(Todas as Disciplinas)"
            self.combo_filtro.Items.Add(item_todas)
            for s in siglas:
                item = ComboBoxItem()
                item.Content = s
                self.combo_filtro.Items.Add(item)
            self.combo_filtro.SelectedIndex = 0
            
        self.atualizar_lista_disponiveis()
        forms.alert("Lista importada com sucesso!\n{} pranchas carregadas.\n{} siglas encontradas.".format(len(self.banco_excel), len(siglas)))

    def _criar_linha_config(self, sigla):
        g = Grid()
        g.Margin = Thickness(0, 0, 0, 5)
        g.ColumnDefinitions.Add(self._col(GridLength(80)))
        g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g.ColumnDefinitions.Add(self._col(GridLength(120)))

        lbl = self._lbl(None, sigla, 12, bold=True, mg=(0, 4, 8, 0))
        Grid.SetColumn(lbl, 0)
        g.Children.Add(lbl)

        combo = ComboBox()
        combo.FontSize = 11
        
        nomes_ordenados = sorted(self.templates_dict.keys())
        lista_ids_para_combo = []
        for nome in nomes_ordenados:
            v_obj = self.templates_dict[nome]
            item = ComboBoxItem()
            item.Content = nome
            combo.Items.Add(item)
            lista_ids_para_combo.append(v_obj.Id if v_obj else None)
        
        self._mapeamento_ids[combo] = lista_ids_para_combo
        combo.SelectedIndex = 0
        Grid.SetColumn(combo, 1)
        g.Children.Add(combo)

        btn_legs = self._btn("Legendas (0)", self.COR_CINZA, size=10)
        btn_legs.Margin = Thickness(5, 0, 0, 0)
        
        def make_on_legs(s_sigla):
            def handler(sender, args):
                opcoes = sorted(self.legends_dict.keys())
                sel = forms.SelectFromList.show(opcoes, title="Legendas para " + s_sigla, multiselect=True)
                if sel is not None:
                    self.legendas_por_sigla[s_sigla] = sel
                    sender.Content = "Legendas ({})".format(len(sel))
            return handler
            
        btn_legs.Click += make_on_legs(sigla)
        Grid.SetColumn(btn_legs, 2)
        g.Children.Add(btn_legs)

        self.painel_config.Children.Add(g)
        self.linhas_config.append((sigla, combo))

    def _get_mapa_config(self):
        mapa = {}
        for sigla, combo in self.linhas_config:
            idx = combo.SelectedIndex
            if idx < 0: continue
            t_id = None
            if combo in self._mapeamento_ids:
                lista = self._mapeamento_ids[combo]
                if idx < len(lista): t_id = lista[idx]
            if t_id is not None: mapa[sigla] = t_id
        return mapa

    def _build_tab_pranchas(self):
        tab = TabItem()
        tab.Header = "Pranchas"
        
        main_grid = Grid()
        main_grid.Margin = Thickness(12)
        
        main_grid.RowDefinitions.Add(self._row(GridLength.Auto))
        main_grid.RowDefinitions.Add(self._row(GridLength.Auto)) 
        main_grid.RowDefinitions.Add(self._row(GridLength.Auto)) 
        main_grid.RowDefinitions.Add(self._row(GridLength(1, GridUnitType.Star))) 
        main_grid.RowDefinitions.Add(self._row(GridLength.Auto)) 
        main_grid.RowDefinitions.Add(self._row(GridLength.Auto))
        main_grid.RowDefinitions.Add(self._row(GridLength.Auto)) 
        main_grid.RowDefinitions.Add(self._row(GridLength(1, GridUnitType.Star))) 

        g_busca = Grid()
        g_busca.Margin = Thickness(0, 0, 0, 8)
        g_busca.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g_busca.ColumnDefinitions.Add(self._col(GridLength.Auto))
        g_busca.ColumnDefinitions.Add(self._col(GridLength.Auto))

        info = StackPanel()
        self._lbl(info, "Adicionar pelo Número da Prancha", 12, bold=True, mg=(0,0,0,2))
        self._lbl(info, "Digite o numero e adicione manualmente se nao quiser usar os filtros.", 10, cor=self.COR_CINZA)
        Grid.SetColumn(info, 0)
        g_busca.Children.Add(info)

        self.txt_busca = TextBox()
        self.txt_busca.FontSize = 12
        self.txt_busca.Padding = Thickness(6, 4, 6, 4)
        self.txt_busca.Width = 120
        self.txt_busca.LostFocus += self.on_buscar_lost_focus
        Grid.SetColumn(self.txt_busca, 1)
        g_busca.Children.Add(self.txt_busca)

        btn_buscar = self._btn("Adicionar >", self.COR_AZUL, size=11, bold=True)
        btn_buscar.Margin = Thickness(6, 0, 0, 0)
        btn_buscar.Click += self.on_buscar_click
        Grid.SetColumn(btn_buscar, 2)
        g_busca.Children.Add(btn_buscar)

        Grid.SetRow(g_busca, 0)
        main_grid.Children.Add(g_busca)

        g_filtro = Grid()
        g_filtro.Margin = Thickness(0, 0, 0, 10)
        g_filtro.ColumnDefinitions.Add(self._col(GridLength.Auto))
        g_filtro.ColumnDefinitions.Add(self._col(GridLength(150)))
        g_filtro.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g_filtro.ColumnDefinitions.Add(self._col(GridLength.Auto))

        lbl_filtro = self._lbl(None, "Filtrar por Disciplina (Sigla):", 11, bold=True, mg=(0,4,8,0))
        Grid.SetColumn(lbl_filtro, 0)
        g_filtro.Children.Add(lbl_filtro)

        self.combo_filtro = ComboBox()
        self.combo_filtro.FontSize = 11
        item_todas = ComboBoxItem()
        item_todas.Content = "(Todas as Disciplinas)"
        self.combo_filtro.Items.Add(item_todas)
        self.combo_filtro.SelectedIndex = 0
        self.combo_filtro.SelectionChanged += self.on_filtro_changed
        Grid.SetColumn(self.combo_filtro, 1)
        g_filtro.Children.Add(self.combo_filtro)

        btn_add_todos = self._btn("↓ Adicionar Todos Listados Abaixo", self.COR_VERDE, size=11, bold=True)
        btn_add_todos.Click += self.on_add_todos_click
        Grid.SetColumn(btn_add_todos, 3)
        g_filtro.Children.Add(btn_add_todos) 

        Grid.SetRow(g_filtro, 1)
        main_grid.Children.Add(g_filtro)

        self._lbl(main_grid, "Pranchas Disponíveis no Excel (Não Selecionadas)", 11, bold=True, cor=self.COR_CINZA, mg=(0,0,0,5))
        Grid.SetRow(main_grid.Children[main_grid.Children.Count - 1], 2)

        scroll_disp = ScrollViewer()
        scroll_disp.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        scroll_disp.Margin = Thickness(0, 0, 0, 10)
        
        self.painel_disponiveis = StackPanel()
        scroll_disp.Content = self.painel_disponiveis
        Grid.SetRow(scroll_disp, 3)
        main_grid.Children.Add(scroll_disp)

        sep = Separator()
        sep.Margin = Thickness(0, 5, 0, 10)
        Grid.SetRow(sep, 4)
        main_grid.Children.Add(sep)

        g_header_criar_title = Grid()
        g_header_criar_title.Margin = Thickness(0, 0, 0, 5)
        g_header_criar_title.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g_header_criar_title.ColumnDefinitions.Add(self._col(GridLength.Auto))

        lbl_title = self._lbl(None, "Pranchas a Serem Criadas (Prontas para o Revit)", 11, bold=True, cor=self.COR_AZUL)
        Grid.SetColumn(lbl_title, 0)
        g_header_criar_title.Children.Add(lbl_title)

        btn_limpar = self._btn("⭯ Limpar Fila", self.COR_VERMELHO, size=10, bold=True)
        btn_limpar.Click += self.on_limpar_tudo_click
        Grid.SetColumn(btn_limpar, 1)
        g_header_criar_title.Children.Add(btn_limpar)

        Grid.SetRow(g_header_criar_title, 5)
        main_grid.Children.Add(g_header_criar_title)

        header_criar = Grid()
        header_criar.Margin = Thickness(0, 0, 0, 5)
        header_criar.ColumnDefinitions.Add(self._col(GridLength(60)))  
        header_criar.ColumnDefinitions.Add(self._col(GridLength(40)))  
        header_criar.ColumnDefinitions.Add(self._col(GridLength(160))) 
        header_criar.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star))) 
        header_criar.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star))) 
        header_criar.ColumnDefinitions.Add(self._col(GridLength(100))) 
        header_criar.ColumnDefinitions.Add(self._col(GridLength(30)))  
        
        colunas = ["Número", "Sigla", "Formato da Prancha", "Vista Base do Revit", "Nome da Prancha", "Legendas", ""]
        for col_idx, txt in enumerate(colunas):
            lbl = self._lbl(None, txt, 10, bold=True, cor=self.COR_CINZA, mg=(0 if col_idx == 0 else 6, 0, 0, 0))
            Grid.SetColumn(lbl, col_idx)
            header_criar.Children.Add(lbl)
        
        Grid.SetRow(header_criar, 6)
        main_grid.Children.Add(header_criar)

        scroll_criar = ScrollViewer()
        scroll_criar.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
        
        self.painel_pranchas = StackPanel()
        scroll_criar.Content = self.painel_pranchas
        Grid.SetRow(scroll_criar, 7)
        main_grid.Children.Add(scroll_criar)

        tab.Content = main_grid
        self.tabs.Items.Add(tab)

    def on_filtro_changed(self, sender, args):
        self.atualizar_lista_disponiveis()

    def on_add_todos_click(self, sender, args):
        if not self.banco_excel: return
        filtro = ""
        if hasattr(self, 'combo_filtro') and self.combo_filtro.SelectedItem:
            filtro = self.combo_filtro.SelectedItem.Content
            if filtro == "(Todas as Disciplinas)": filtro = ""

        numeros_ordenados = sorted(self.banco_excel.keys())
        for num in numeros_ordenados:
            if num not in self.pranchas_selecionadas:
                dados = self.banco_excel[num]
                if filtro == "" or dados["sigla"] == filtro:
                    self.pranchas_selecionadas.add(num)
                    legendas_base = self.legendas_por_sigla.get(dados["sigla"], [])
                    self._add_linha_prancha_criacao(
                        numero=dados["numero"],
                        nome_arquivo=dados["nome_arquivo"],
                        sigla=dados["sigla"],
                        descricao=dados["descricao"],
                        legendas_iniciais=legendas_base
                    )
        self.atualizar_lista_disponiveis()

    def on_limpar_tudo_click(self, sender, args):
        self.pranchas_selecionadas.clear()
        self.painel_pranchas.Children.Clear()
        self.campos_pranchas = []
        self.atualizar_lista_disponiveis()

    def atualizar_lista_disponiveis(self):
        self.painel_disponiveis.Children.Clear()
        if not self.banco_excel:
            self._lbl(self.painel_disponiveis, "Nenhum Excel importado ainda.", 11, cor=self.COR_CINZA, mg=(10,0,0,0))
            return

        filtro = ""
        if hasattr(self, 'combo_filtro') and self.combo_filtro.SelectedItem:
            filtro = self.combo_filtro.SelectedItem.Content
            if filtro == "(Todas as Disciplinas)": filtro = ""

        numeros_ordenados = sorted(self.banco_excel.keys())
        adicionadas = 0
        
        for num in numeros_ordenados:
            if num not in self.pranchas_selecionadas:
                dados = self.banco_excel[num]
                if filtro == "" or dados["sigla"] == filtro:
                    self._add_linha_disponivel(dados)
                    adicionadas += 1
                
        if adicionadas == 0:
            if filtro:
                self._lbl(self.painel_disponiveis, "Todas as pranchas da disciplina {} ja foram adicionadas.".format(filtro), 11, cor=self.COR_CINZA, mg=(10,0,0,0))
            else:
                self._lbl(self.painel_disponiveis, "Todas as pranchas ja foram selecionadas para criacao.", 11, cor=self.COR_CINZA, mg=(10,0,0,0))

    def _add_linha_disponivel(self, dados):
        g = Grid()
        g.Margin = Thickness(0, 0, 0, 3)
        g.ColumnDefinitions.Add(self._col(GridLength(65)))
        g.ColumnDefinitions.Add(self._col(GridLength(50)))
        g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g.ColumnDefinitions.Add(self._col(GridLength(35)))

        lbl_num = self._lbl(None, dados["numero"], 11, bold=True)
        Grid.SetColumn(lbl_num, 0)
        g.Children.Add(lbl_num)

        lbl_sigla = self._lbl(None, dados["sigla"], 11, cor=self.COR_CINZA, mg=(5,0,0,0))
        Grid.SetColumn(lbl_sigla, 1)
        g.Children.Add(lbl_sigla)

        texto_exibicao = "{} - {}".format(dados["nome_arquivo"], dados["descricao"])
        lbl_nome = self._lbl(None, texto_exibicao, 11, mg=(5,0,0,0))
        Grid.SetColumn(lbl_nome, 2)
        g.Children.Add(lbl_nome)

        btn_add = Button()
        btn_add.Content = " + "
        btn_add.FontSize = 14
        btn_add.FontWeight = FontWeights.Bold
        btn_add.Foreground = SolidColorBrush(self.COR_BRANCO)
        btn_add.Background = SolidColorBrush(self.COR_VERDE)
        btn_add.Padding = Thickness(0)
        
        def make_mover(num):
            def handler(sender, args):
                self.mover_para_criacao(num)
            return handler
            
        btn_add.Click += make_mover(dados["numero"])
        Grid.SetColumn(btn_add, 3)
        g.Children.Add(btn_add)

        self.painel_disponiveis.Children.Add(g)

    def on_buscar_lost_focus(self, sender, args):
        self.executar_busca()

    def on_buscar_click(self, sender, args):
        self.executar_busca()

    def executar_busca(self):
        buscado = self.txt_busca.Text.strip()
        if not buscado: return
        if not self.banco_excel:
            forms.alert("Por favor, importe a Lista Mestra na aba Configuracoes primeiro.")
            self.txt_busca.Text = ""
            return
            
        match_key = None
        if buscado in self.banco_excel:
            match_key = buscado
        else:
            for k, v in self.banco_excel.items():
                if v.get("numero_original") == buscado:
                    match_key = k
                    break

        if not match_key:
            forms.alert("Numero '{}' nao encontrado no Excel importado.".format(buscado))
            self.txt_busca.Text = ""
            return
            
        self.mover_para_criacao(match_key)
        self.txt_busca.Text = ""

    def mover_para_criacao(self, numero):
        if numero in self.pranchas_selecionadas: return 
        self.pranchas_selecionadas.add(numero)
        dados = self.banco_excel[numero]
        legendas_base = self.legendas_por_sigla.get(dados["sigla"], [])
        self._add_linha_prancha_criacao(
            numero=dados["numero"],
            nome_arquivo=dados["nome_arquivo"],
            sigla=dados["sigla"],
            descricao=dados["descricao"],
            legendas_iniciais=legendas_base
        )
        self.atualizar_lista_disponiveis()

    def _add_linha_prancha_criacao(self, numero="", nome_arquivo="", sigla="", descricao="", vista_ref=None, legendas_iniciais=None, manual_template_id=None, is_manual=False, manual_base_view_str=None, manual_tb_str=None):
        if legendas_iniciais is None:
            legendas_iniciais = []

        g = Grid()
        g.Margin = Thickness(0, 0, 0, 5)
        g.ColumnDefinitions.Add(self._col(GridLength(60)))
        g.ColumnDefinitions.Add(self._col(GridLength(40)))
        g.ColumnDefinitions.Add(self._col(GridLength(160)))
        g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))
        g.ColumnDefinitions.Add(self._col(GridLength(100)))
        g.ColumnDefinitions.Add(self._col(GridLength(30))) 

        txt_num = TextBox()
        txt_num.FontSize = 11
        txt_num.Padding = Thickness(5, 4, 5, 4)
        txt_num.Text = numero
        txt_num.IsReadOnly = True
        txt_num.Background = SolidColorBrush(Color.FromRgb(240, 240, 240))
        Grid.SetColumn(txt_num, 0)
        g.Children.Add(txt_num)

        txt_sigla = TextBox()
        txt_sigla.FontSize = 11
        txt_sigla.Padding = Thickness(5, 4, 5, 4)
        txt_sigla.Margin = Thickness(5, 0, 5, 0)
        txt_sigla.Text = sigla
        txt_sigla.IsReadOnly = True
        txt_sigla.Background = SolidColorBrush(Color.FromRgb(240, 240, 240))
        Grid.SetColumn(txt_sigla, 1)
        g.Children.Add(txt_sigla)

        combo_tb = ComboBox()
        combo_tb.FontSize = 11
        combo_tb.Margin = Thickness(0, 0, 5, 0)
        sorted_tbs = sorted(self.titleblocks_dict.keys())
        for tb_name in sorted_tbs:
            item = ComboBoxItem()
            item.Content = tb_name
            combo_tb.Items.Add(item)
            
        best_tb_idx = 0
        if manual_tb_str and manual_tb_str in sorted_tbs:
            best_tb_idx = sorted_tbs.index(manual_tb_str)
            
        if combo_tb.Items.Count > 0:
            combo_tb.SelectedIndex = best_tb_idx
            
        Grid.SetColumn(combo_tb, 2)
        g.Children.Add(combo_tb)

        combo_base = ComboBox()
        combo_base.FontSize = 11
        combo_base.Margin = Thickness(0, 0, 5, 0)
        
        sorted_views = sorted(self.views_dict.keys())
        for v_name in sorted_views:
            item = ComboBoxItem()
            item.Content = v_name
            combo_base.Items.Add(item)

        best_match_idx = -1
        if manual_base_view_str:
            for idx, v_name in enumerate(sorted_views):
                if v_name == manual_base_view_str:
                    best_match_idx = idx
                    break
        elif descricao:
            desc_upper = remove_accents(descricao).upper()
            for idx, v_name in enumerate(sorted_views):
                v_clean = remove_accents(v_name).upper()
                v_real = v_clean.split("]")[-1].strip()
                if v_real in desc_upper:
                    best_match_idx = idx
                    break
        
        if combo_base.Items.Count > 0:
            combo_base.SelectedIndex = best_match_idx if best_match_idx != -1 else 0

        Grid.SetColumn(combo_base, 3)
        g.Children.Add(combo_base)

        txt_nome = TextBox()
        txt_nome.FontSize = 11
        txt_nome.Padding = Thickness(5, 4, 5, 4)
        txt_nome.Text = nome_arquivo
        Grid.SetColumn(txt_nome, 4)
        g.Children.Add(txt_nome)

        wrap_legs = WrapPanel()
        wrap_legs.Margin = Thickness(5, 0, 5, 0)
        Grid.SetColumn(wrap_legs, 5)
        g.Children.Add(wrap_legs)

        lista_legendas_ativas = list(legendas_iniciais)

        def renderizar_chips():
            wrap_legs.Children.Clear()
            if not lista_legendas_ativas:
                self._lbl(wrap_legs, "(Nenhuma)", 10, cor=self.COR_CINZA)
                return
            for leg_nome in lista_legendas_ativas:
                chip = Grid()
                chip.Margin = Thickness(0, 0, 4, 4)
                chip.Background = SolidColorBrush(Color.FromRgb(220, 230, 240))
                chip.ColumnDefinitions.Add(self._col(GridLength.Auto))
                chip.ColumnDefinitions.Add(self._col(GridLength.Auto))

                l = self._lbl(None, leg_nome[:10] + "..." if len(leg_nome)>12 else leg_nome, 9)
                l.Padding = Thickness(4, 2, 2, 2)
                Grid.SetColumn(l, 0)
                chip.Children.Add(l)

                b = Button()
                b.Content = "x"
                b.FontSize = 9
                b.Background = SolidColorBrush(Color.FromRgb(220, 230, 240))
                b.BorderThickness = Thickness(0)
                b.Foreground = SolidColorBrush(Color.FromRgb(200, 50, 50))
                b.Padding = Thickness(2, 0, 4, 0)

                def make_rem_leg(n):
                    def handler(s, a):
                        if n in lista_legendas_ativas: lista_legendas_ativas.remove(n)
                        renderizar_chips()
                    return handler
                b.Click += make_rem_leg(leg_nome)

                Grid.SetColumn(b, 1)
                chip.Children.Add(b)
                wrap_legs.Children.Add(chip)

        renderizar_chips()

        campo_dict = {
            "txt_num":          txt_num,
            "txt_sigla":        txt_sigla,
            "combo_tb":         combo_tb,
            "combo_base":       combo_base,
            "txt_nome":         txt_nome,
            "vista_ref":        vista_ref,
            "descricao_oculta": descricao,
            "legendas_ativas":  lista_legendas_ativas,
            "is_manual":        is_manual,
            "manual_template_id": manual_template_id
        }

        btn_del = Button()
        btn_del.Content = "X"
        btn_del.FontSize = 10
        btn_del.FontWeight = FontWeights.Bold
        btn_del.Foreground = SolidColorBrush(self.COR_BRANCO)
        btn_del.Background = SolidColorBrush(Color.FromRgb(220, 50, 50))
        btn_del.BorderThickness = Thickness(0)
        btn_del.Margin = Thickness(5, 0, 0, 0)
        Grid.SetColumn(btn_del, 6)

        def make_remover(linha_grid, d_campo, num):
            def handler(sender, args):
                self.painel_pranchas.Children.Remove(linha_grid)
                if d_campo in self.campos_pranchas: self.campos_pranchas.remove(d_campo)
                if num in self.pranchas_selecionadas: self.pranchas_selecionadas.remove(num)
                self.atualizar_lista_disponiveis()
            return handler

        btn_del.Click += make_remover(g, campo_dict, numero)
        g.Children.Add(btn_del)

        self.painel_pranchas.Children.Add(g)
        self.campos_pranchas.append(campo_dict)

    def _build_tab_carimbo(self):
        tab = TabItem()
        tab.Header = "Carimbo"
        outer = StackPanel()
        outer.Margin = Thickness(12, 12, 12, 12)

        self._lbl(outer, "Parametros do carimbo", 12, bold=True, mg=(0,0,0,4))
        self._lbl(outer, "Estes valores serao aplicados em TODAS as pranchas criadas.\nNumero e Nome sao definidos na aba Pranchas.", 10, cor=self.COR_CINZA, mg=(0,0,0,10))
        outer.Children.Add(self._sep(8))

        self.campos_carimbo = {}

        for param_name, label_text in PARAMS_CARIMBO:
            g = Grid()
            g.Margin = Thickness(0, 0, 0, 6)
            g.ColumnDefinitions.Add(self._col(GridLength(150)))
            g.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))

            lbl = self._lbl(None, param_name, 11, mg=(0, 4, 8, 0))
            Grid.SetColumn(lbl, 0)
            g.Children.Add(lbl)

            txt = TextBox()
            txt.FontSize = 11
            txt.Padding = Thickness(6, 4, 6, 4)
            Grid.SetColumn(txt, 1)
            g.Children.Add(txt)

            outer.Children.Add(g)
            self.campos_carimbo[label_text] = txt
        
        outer.Children.Add(self._sep(8))
        self._lbl(outer, "Tipo de Viewport (Titulo da Vista Principal)", 12, bold=True, mg=(0,0,0,6))

        g_vp = Grid()
        g_vp.Margin = Thickness(0, 0, 0, 6)
        g_vp.ColumnDefinitions.Add(self._col(GridLength(150)))
        g_vp.ColumnDefinitions.Add(self._col(GridLength(1, GridUnitType.Star)))

        lbl_vp = self._lbl(None, "Tipo de Viewport:", 11, mg=(0, 4, 8, 0))
        Grid.SetColumn(lbl_vp, 0)
        g_vp.Children.Add(lbl_vp)

        self.combo_vp_type = ComboBox()
        self.combo_vp_type.FontSize = 11
        self._vp_type_ids_list = []

        nomes_ordenados = sorted(self.viewport_types_dict.keys())
        for nome in nomes_ordenados:
            item = ComboBoxItem()
            item.Content = nome
            self.combo_vp_type.Items.Add(item)
            self._vp_type_ids_list.append(self.viewport_types_dict[nome])

        if self.combo_vp_type.Items.Count > 0:
            self.combo_vp_type.SelectedIndex = 0

        Grid.SetColumn(self.combo_vp_type, 1)
        g_vp.Children.Add(self.combo_vp_type)
        outer.Children.Add(g_vp)

        outer.Children.Add(self._sep(8))
        self._lbl(outer, "ATENCAO: os nomes dos parametros acima devem coincidir exatamente\ncom os parametros da familia do carimbo no Revit.", 9, cor=self.COR_CINZA, mg=(0,0,0,0))

        tab.Content = outer
        self.tabs.Items.Add(tab)

    def on_criar(self, sender, args):
        mapa_cfg = self._get_mapa_config()
        pranchas = []

        for campo in self.campos_pranchas:
            numero = campo["txt_num"].Text.strip()
            nome   = campo["txt_nome"].Text.strip()
            sigla  = campo["txt_sigla"].Text.strip().upper()
            
            if not numero and not nome: continue

            tb_label = campo["combo_tb"].SelectedItem.Content if campo["combo_tb"].SelectedItem else None
            tb_symbol = self.titleblocks_dict.get(tb_label)
            
            if not tb_symbol:
                forms.alert("Linha {}: Selecione um formato de prancha válido.".format(numero))
                return

            base_view_label = campo["combo_base"].SelectedItem.Content if campo["combo_base"].SelectedItem else None
            base_view = self.views_dict.get(base_view_label)

            if base_view is None:
                forms.alert("Linha {}: Selecione uma Vista Base valida.".format(numero))
                return

            if campo.get("is_manual"): template_id = campo.get("manual_template_id")
            else: template_id = mapa_cfg.get(sigla)
            
            leg_nomes = campo.get("legendas_ativas", [])
            leg_objs = [self.legends_dict[n] for n in self.legends_dict if n in leg_nomes]

            pranchas.append({
                "numero":                 numero,
                "nome_arquivo":           nome,
                "descricao":              campo.get("descricao_oculta", ""),
                "sigla":                  sigla,
                "tb_symbol":              tb_symbol,
                "template_id":            template_id,
                "base_view":              base_view,
                "vista_ref":              campo["vista_ref"],
                "legendas_desta_prancha": leg_objs
            })

        if not pranchas:
            forms.alert("Nenhuma prancha configurada para criar.")
            return

        dados_carimbo = {}
        for param_name, txt in self.campos_carimbo.items():
            valor = txt.Text.strip()
            if valor: dados_carimbo[param_name] = valor

        idx_vp = self.combo_vp_type.SelectedIndex
        vp_type_id_selecionado = None
        if idx_vp >= 0 and idx_vp < len(self._vp_type_ids_list):
            vp_type_id_selecionado = self._vp_type_ids_list[idx_vp]

        self.resultado = {
            "pranchas":              pranchas,
            "dados_carimbo":         dados_carimbo,
            "vp_type_id_selecionado": vp_type_id_selecionado
        }
        self.Close()

    def _lbl(self, parent, text, size, bold=False, cor=None, mg=None):
        lbl = Label()
        lbl.Content = text
        lbl.FontSize = size
        if bold: lbl.FontWeight = FontWeights.Bold
        if cor: lbl.Foreground = SolidColorBrush(cor)
        if mg: lbl.Margin = Thickness(*mg) if len(mg) > 1 else Thickness(mg[0])
        if parent is not None and hasattr(parent, "Children"): parent.Children.Add(lbl)
        return lbl

    def _sep(self, bottom=8):
        s = Separator()
        s.Margin = Thickness(0, 0, 0, bottom)
        return s

    def _btn(self, text, cor_fundo, size=12, bold=False):
        b = Button()
        b.Content = text
        b.FontSize = size
        b.Padding = Thickness(10, 5, 10, 5)
        b.Background = SolidColorBrush(cor_fundo)
        b.Foreground = SolidColorBrush(self.COR_BRANCO)
        if bold: b.FontWeight = FontWeights.Bold
        return b

def executar_fluxo_excel_wpf():
    views_dict = get_all_views()
    if not views_dict:
        forms.alert("Nenhuma vista duplicavel encontrada no projeto.", exitscript=True)

    templates_dict     = get_all_view_templates()
    legends_dict       = get_all_legends()
    viewport_types_dict = get_all_viewport_types()
    titleblocks_dict   = get_titleblocks_dict()

    if not titleblocks_dict:
        forms.alert("Nenhum carimbo encontrado no projeto.", exitscript=True)
    
    janela = CriarPranchasWindow(views_dict, templates_dict, legends_dict, viewport_types_dict, titleblocks_dict)
    janela.ShowDialog()

    if not janela.resultado:
        script.exit()

    dados                    = janela.resultado
    pranchas                 = dados["pranchas"]
    dados_carimbo            = dados["dados_carimbo"]
    target_vp_type_id        = dados.get("vp_type_id_selecionado") or ElementId.InvalidElementId

    todos_tbs = get_todos_titleblocks_com_area_util()

    with Transaction(doc, "Ativar Carimbos") as t:
        t.Start()
        for sym in titleblocks_dict.values():
            if not sym.IsActive:
                try: sym.Activate()
                except: pass
        t.Commit()

    criadas = []
    erros   = []
    avisos  = []
    vistas_para_template = []
    vps_para_centralizar = []

    todas_pranchas = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
    numeros_existentes = set([p.SheetNumber for p in todas_pranchas])

    with Transaction(doc, "Criar Pranchas e Vistas") as t:
        t.Start()

        for item in pranchas:
            numero_base = item["numero"]
            nome_arquivo= item["nome_arquivo"]
            descricao   = item["descricao"]
            template_id = item["template_id"]
            sigla       = item["sigla"]
            base_view   = item.get("base_view")
            tb_symbol   = item["tb_symbol"]

            numero = numero_base
            while numero in numeros_existentes: numero += "*"
            numeros_existentes.add(numero)

            try:
                nova_vista_id = base_view.Duplicate(ViewDuplicateOption.WithDetailing)
                nova_vista    = doc.GetElement(nova_vista_id)

                nome_base_origem = base_view.Name if base_view else "Detalhe"
                novo_nome_base = "{} - {}".format(nome_base_origem, sigla) if sigla else "{} - Detalhe".format(nome_base_origem)
                novo_nome = novo_nome_base
                contador = 1
                
                while True:
                    try:
                        nova_vista.Name = novo_nome
                        break
                    except Exceptions.ArgumentException:
                        novo_nome = "{} ({})".format(novo_nome_base, contador)
                        contador += 1

                if template_id:
                    try:
                        nova_vista.ViewTemplateId = template_id
                        doc.Regenerate()
                    except Exception:
                        pass 
                    vistas_para_template.append((nova_vista, template_id))

                modelo_w = 0.0
                modelo_h = 0.0
                
                try:
                    outline = nova_vista.Outline
                    escala_atual = nova_vista.Scale
                    
                    if escala_atual <= 0:
                        escala_atual = 100
                        
                    modelo_w = abs(outline.Max.U - outline.Min.U) * escala_atual
                    modelo_h = abs(outline.Max.V - outline.Min.V) * escala_atual
                    
                except Exception:
                    if nova_vista.CropBoxActive:
                        bb_crop = nova_vista.CropBox
                        modelo_w = abs(bb_crop.Max.X - bb_crop.Min.X)
                        modelo_h = abs(bb_crop.Max.Y - bb_crop.Min.Y)

                if not tb_symbol: tb_symbol = todos_tbs[-1][0]
                sheet = ViewSheet.Create(doc, tb_symbol.Id)
                sheet.SheetNumber = numero 
                sheet.Name = nome_arquivo
                doc.Regenerate()
                
                carimbos = FilteredElementCollector(doc, sheet.Id).OfCategory(BuiltInCategory.OST_TitleBlocks).ToElements()
                tb_instance = carimbos[0] if carimbos else None

                legendas_desta_prancha = item.get("legendas_desta_prancha", [])
                legend_data = []
                w_leg_max = 0
                h_leg_total = 135.0 / 304.8 
                
                if legendas_desta_prancha:
                    for leg in legendas_desta_prancha:
                        if leg is None: continue
                        try:
                            if not Viewport.CanAddViewToSheet(doc, sheet.Id, leg.Id):
                                erros.append(u"Legenda '{}' não suportada.".format(leg.Name))
                                continue
                            vp_leg = Viewport.Create(doc, sheet.Id, leg.Id, XYZ(0, 0, 0))
                            doc.Regenerate()
                            out = vp_leg.GetBoxOutline()
                            w = out.MaximumPoint.X - out.MinimumPoint.X
                            h = out.MaximumPoint.Y - out.MinimumPoint.Y
                            legend_data.append((vp_leg, w, h))
                            if w > w_leg_max: w_leg_max = w
                            h_leg_total += h + (2.0 / 304.8)
                        except Exception as e_leg:
                            erros.append(u"Erro ao medir legenda '{}': {}".format(getattr(leg, 'Name', '?'), str(e_leg)))

                escala_escolhida = ESCALAS_VALIDAS[-1]
                aviso = None

                if modelo_w > 0 and modelo_h > 0:
                    melhor_tb, escala_escolhida, aviso = escolher_melhor_titleblock(
                        modelo_w, modelo_h,
                        tb_symbol,
                        todos_tbs,
                        nova_vista.Name,
                        w_leg_max,
                        h_leg_total
                    )
                    if tb_instance and melhor_tb.Id != tb_instance.GetTypeId():
                        tb_instance.ChangeTypeId(melhor_tb.Id)
                        doc.Regenerate()
                else:
                    if modelo_w == 0 or modelo_h == 0:
                        aviso = "AVISO: Não foi possível medir a vista '{}'. Usando formato padrão.".format(nova_vista.Name)

                if aviso: avisos.append(aviso)

                try:
                    orig_template_pre = nova_vista.ViewTemplateId
                    if orig_template_pre != DB.ElementId.InvalidElementId: nova_vista.ViewTemplateId = DB.ElementId.InvalidElementId
                    nova_vista.Scale = escala_escolhida
                    if orig_template_pre != DB.ElementId.InvalidElementId: nova_vista.ViewTemplateId = orig_template_pre
                except: pass

                for param_name, valor in dados_carimbo.items():
                    p = sheet.LookupParameter(param_name)
                    if not p and tb_instance: p = tb_instance.LookupParameter(param_name)
                    if not p: p = doc.ProjectInformation.LookupParameter(param_name)
                    
                    if p and not p.IsReadOnly:
                        if p.StorageType == StorageType.String: p.Set(str(valor))
                        elif p.StorageType == StorageType.Integer:
                            try: p.Set(int(valor))
                            except: pass
                        elif p.StorageType == StorageType.Double:
                            try: p.Set(float(valor))
                            except: pass
                
                set_param_titulo(sheet, tb_instance, descricao)
                
                if tb_instance and sigla:
                    target_param = SIGLA_TO_PARAM.get(sigla)
                    for p in tb_instance.Parameters:
                        if p.IsReadOnly or p.StorageType != StorageType.Integer: continue
                        p_name = p.Definition.Name
                        p_name_clean = remove_accents(p_name).upper()
                        
                        if "PROJETO" in p_name_clean or "DISCIPLINA" in p_name_clean:
                            try: p.Set(0)
                            except: pass
                            
                            if target_param and p_name == target_param:
                                try: p.Set(1)
                                except: pass
                            elif not target_param:
                                match_terms = PARAM_DISC_MATCH.get(sigla, [sigla])
                                palavras_param = p_name_clean.replace("-", " ").replace("/", " ").split()
                                match = False
                                for term in match_terms:
                                    if term in p_name_clean:
                                        if len(term) <= 2:
                                            if term in palavras_param: match = True
                                        else: match = True
                                if match:
                                    try: p.Set(1)
                                    except: pass

                if tb_instance and legend_data:
                    doc.Regenerate()
                    bb_tb = tb_instance.get_BoundingBox(sheet)
                    if bb_tb:
                        margem_direita_mm = 25.0 
                        altura_selo_mm = 135.0    
                        target_x_max = bb_tb.Max.X - (margem_direita_mm / 304.8)
                        current_y_min = bb_tb.Min.Y + (altura_selo_mm / 304.8)
                        
                        for vp_leg, w_vp, h_vp in legend_data:
                            novo_x = target_x_max - (w_vp / 2.0)
                            novo_y = current_y_min + (h_vp / 2.0)
                            vp_leg.SetBoxCenter(XYZ(novo_x, novo_y, 0))
                            current_y_min += h_vp + (2.0 / 304.8)
                        doc.Regenerate()

                if Viewport.CanAddViewToSheet(doc, sheet.Id, nova_vista.Id):
                    vp = Viewport.Create(doc, sheet.Id, nova_vista.Id, XYZ(1.0, 0.7, 0))
                    if target_vp_type_id: vp.ChangeTypeId(target_vp_type_id)
                    set_viewport_titulo(vp, nova_vista.Name)
                    vps_para_centralizar.append((sheet.Id, vp.Id))
                    criadas.append("{} - {} [{}]".format(numero, nome_arquivo, sigla or "sem sigla"))
            except Exception as e:
                import traceback
                erros.append("{} - {} > {}\n{}".format(numero, nome_arquivo, str(e), traceback.format_exc()))

        t.Commit()

    if vistas_para_template:
        with Transaction(doc, "Aplicar View Templates") as t:
            t.Start()
            for v, tid in vistas_para_template:
                try: v.ViewTemplateId = tid
                except: pass
            t.Commit()

    if vps_para_centralizar:
        with Transaction(doc, "Otimizar Pranchas e Escala") as t:
            t.Start()
            mm_to_ft = 1.0 / 304.8
            margem_esq_pes = MARGEM_ESQ_MM * mm_to_ft
            margem_sup_pes = MARGEM_SUP_MM * mm_to_ft
            margem_inf_pes = MARGEM_INF_MM * mm_to_ft

            for sheet_id, vp_id in vps_para_centralizar:
                try:
                    sheet_elem = doc.GetElement(sheet_id)
                    vp_elem = doc.GetElement(vp_id)
                    nova_vista = doc.GetElement(vp_elem.ViewId)
                    
                    tbs = FilteredElementCollector(doc, sheet_id).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
                    if not tbs: continue
                    
                    tb_inst = tbs[0]
                    bb = tb_inst.get_BoundingBox(sheet_elem)
                    if not bb: continue
                    
                    util_min_x = bb.Min.X + margem_esq_pes
                    util_max_y = bb.Max.Y - margem_sup_pes
                    util_min_y = bb.Min.Y + margem_inf_pes
                    
                    out_vp = vp_elem.GetBoxOutline()
                    w_vp = out_vp.MaximumPoint.X - out_vp.MinimumPoint.X
                    h_vp = out_vp.MaximumPoint.Y - out_vp.MinimumPoint.Y
                    
                    target_x = util_min_x + (w_vp / 2.0)
                    target_y = util_max_y - (h_vp / 2.0)
                    
                    orig_anno = None
                    try:
                        orig_anno = nova_vista.AreAnnotationCategoriesHidden
                        nova_vista.AreAnnotationCategoriesHidden = True
                    except: pass
                        
                    doc.Regenerate()
                    vp_elem.SetBoxCenter(XYZ(target_x, target_y, 0))
                    
                    if orig_anno is not None:
                        try: nova_vista.AreAnnotationCategoriesHidden = orig_anno
                        except: pass
                except Exception: pass

            t.Commit()

    msg = u"Pranchas criadas: {}\nErros: {}".format(len(criadas), len(erros))
    if erros: msg += u"\n\nDetalhes dos erros:\n{}".format(u"\n".join(erros))
    if avisos: msg += u"\n\nAvisos de encaixe ({}):\n{}".format(len(avisos), u"\n".join(avisos))
    forms.alert(msg)

if __name__ == '__main__':
    executar_fluxo_excel_wpf()