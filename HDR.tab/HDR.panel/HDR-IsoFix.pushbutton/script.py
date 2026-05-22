# -*- coding: utf-8 -*-

__title__ = "Iso\Fix+"
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
    ViewType, Level
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

def get_unique_view_name(base_name):
    existing = {v.Name for v in FilteredElementCollector(doc).OfClass(DB.View)}
    if base_name not in existing:
        return base_name
    i = 1
    while True:
        candidate = "{} ({})".format(base_name, i)
        if candidate not in existing:
            return candidate
        i += 1

def get_titleblock_type():
    for sym in (FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .OfClass(FamilySymbol)):
        if sym.Family.Name == FAMILIA_NOME:
            return sym

    all_types = {}
    for s in (FilteredElementCollector(doc)
              .OfCategory(BuiltInCategory.OST_TitleBlocks)
              .OfClass(FamilySymbol)):
        p = s.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        label = "{} : {}".format(s.Family.Name, p.AsString() if p else "?")
        all_types[label] = s

    if not all_types:
        forms.alert(
            "Nenhum carimbo encontrado.\nCarregue a família '{}' e tente novamente.".format(FAMILIA_NOME),
            exitscript=True
        )

    chosen = forms.SelectFromList.show(
        sorted(all_types.keys()), title="Selecione o Carimbo", multiselect=False
    )
    if not chosen:
        script.exit()
    return all_types[chosen]

def get_3d_view_family_types():
    tipos = {}
    for vft in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if vft.ViewFamily == ViewFamily.ThreeDimensional:
            p = vft.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            nome = p.AsString() if p else vft.Id.ToString()
            tipos[nome] = vft
    return tipos


class IsometricConfigWindow(Window):
    COR_AZUL   = Color.FromRgb(30, 80, 160)
    COR_VERDE  = Color.FromRgb(22, 160, 80)
    COR_CINZA  = Color.FromRgb(120, 120, 120)
    COR_BRANCO = Color.FromRgb(255, 255, 255)
    COR_FUNDO  = Color.FromRgb(245, 245, 245)

    def __init__(self, templates_dict, view_types_dict):
        self.templates_dict  = templates_dict
        self.view_types_dict = view_types_dict
        self.resultado       = None

        self._template_names  = sorted(templates_dict.keys())
        self._view_type_names = sorted(view_types_dict.keys())
        self.campos_carimbo   = {}

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

        self._lbl(pnl, "Prefixo (ex: 'ISOMÉTRICO', 'ISO'):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_prefixo = TextBox()
        self.txt_prefixo.FontSize = 12
        self.txt_prefixo.Padding  = Thickness(6, 4, 6, 4)
        self.txt_prefixo.Margin   = Thickness(0, 0, 0, 12)
        self.txt_prefixo.Text     = "ISOMÉTRICO"
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
        self.txt_num_inicial.Text     = "1"
        pnl_num.Children.Add(self.txt_num_inicial)
        Grid.SetColumn(pnl_num, 0); g_cont.Children.Add(pnl_num)

        pnl_zero = StackPanel(); pnl_zero.Margin = Thickness(4, 0, 0, 0)
        self._lbl(pnl_zero, "Formato do Número:", 11, bold=True, mg=(0, 0, 0, 4))
        self.combo_zeros = ComboBox(); self.combo_zeros.FontSize = 11
        i1 = ComboBoxItem(); i1.Content = "1 dígito (1, 2, 3)"
        i2 = ComboBoxItem(); i2.Content = "2 dígitos (01, 02)"
        i3 = ComboBoxItem(); i3.Content = "3 dígitos (001, 002)"
        self.combo_zeros.Items.Add(i1)
        self.combo_zeros.Items.Add(i2)
        self.combo_zeros.Items.Add(i3)
        self.combo_zeros.SelectedIndex = 1
        pnl_zero.Children.Add(self.combo_zeros)
        Grid.SetColumn(pnl_zero, 1); g_cont.Children.Add(pnl_zero)

        pnl.Children.Add(g_cont)

        self._lbl(pnl, "Sufixo (ex: '- 1°TIPO' ou vazio):", 11, bold=True, mg=(0, 0, 0, 4))
        self.txt_sufixo = TextBox()
        self.txt_sufixo.FontSize = 12
        self.txt_sufixo.Padding  = Thickness(6, 4, 6, 4)
        self.txt_sufixo.Margin   = Thickness(0, 0, 0, 16)
        self.txt_sufixo.Text     = "- 1°TIPO"
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
        self.chk_sem_escala.Content = "Sem Escala (ignora o campo acima)"
        self.chk_sem_escala.FontSize = 11
        self.chk_sem_escala.Margin   = Thickness(0, 0, 0, 12)
        self.chk_sem_escala.IsChecked = True
        pnl.Children.Add(self.chk_sem_escala)

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
        tab.Content = scroll
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
            "dados_carimbo": dados_carimbo,
        }
        self.Close()

def executar_fluxo_isometrico():
    tb_type = get_titleblock_type()
    if not tb_type:
        script.exit()
    tb_type_id = tb_type.Id

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

    dlg = IsometricConfigWindow(templates, view_types_3d)
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
    dados_carimbo = dlg.resultado["dados_carimbo"]

    todos_niveis = sorted(FilteredElementCollector(doc).OfClass(Level).ToElements(), key=lambda n: n.Elevation)
    nivel_atual = doc.ActiveView.GenLevel

    grupos = []          
    contador = 1

    while True:
        formato_num  = "{:0" + str(zeros) + "d}"
        numero_str   = formato_num.format(num_inicial + contador - 1)
        contador_full = "{}{}".format(identificador, numero_str)
        partes        = [prefixo, contador_full, sufixo]
        nome_preview  = " ".join([p for p in partes if p])

        instrucao = "[{}] Desenhe o retângulo definindo a área do isométrico (ESC para opções)".format(
            nome_preview
        )

        try:
            box = uidoc.Selection.PickBox(PickBoxStyle.Directional, instrucao)
            grupos.append(box)
            contador += 1

        except Exceptions.OperationCanceledException:
            if grupos:
                opcao = forms.alert(
                    "Seleção pausada. {} isométrico(s) na fila.\nO que deseja fazer?".format(len(grupos)),
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
            forms.alert("Erro na seleção: {}".format(str(e)))
            break

    if not grupos:
        forms.alert("Nenhum isométrico selecionado. Operação cancelada.")
        script.exit()

    vistas_geradas = []
    erros          = []

    with TransactionGroup(doc, "Gerador de Isométricos") as tg:
        tg.Start()

        formato_num = "{:0" + str(zeros) + "d}"

        for idx, box in enumerate(grupos):
            numero_str    = formato_num.format(num_inicial + idx)
            contador_full = "{}{}".format(identificador, numero_str)
            partes        = [prefixo, contador_full, sufixo]
            nome_iso      = sanitize_name(" ".join([p for p in partes if p]))

            novo_id = None

            with Transaction(doc, "Criar Isométrico: {}".format(nome_iso)) as t:
                t.Start()
                try:
                    p_min_x = min(box.Min.X, box.Max.X)
                    p_max_x = max(box.Min.X, box.Max.X)
                    p_min_y = min(box.Min.Y, box.Max.Y)
                    p_max_y = max(box.Min.Y, box.Max.Y)

                    margem = 0.5 / 0.3048

                    # Limites exatos do Z fixados para 2.30m acima da base
                    if nivel_atual:
                        z_level = nivel_atual.Elevation
                        z_min = z_level - margem
                        z_max = z_level + (2.30 / 0.3048)
                    else:
                        z_base = min(box.Min.Z, box.Max.Z)
                        z_min = z_base - margem
                        z_max = z_base + (2.30 / 0.3048)

                    section_box = DB.BoundingBoxXYZ()
                    section_box.Min = XYZ(p_min_x - margem, p_min_y - margem, z_min)
                    section_box.Max = XYZ(p_max_x + margem, p_max_y + margem, z_max)

                    nova = View3D.CreateIsometric(doc, view_type_id)
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

                    doc.Regenerate()
                    novo_id = nova.Id

                    t.Commit()
                except Exception as e:
                    erros.append("Erro ao criar '{}': {}".format(nome_iso, str(e)))
                    t.RollBack()

            if novo_id is None:
                continue

            vistas_geradas.append({
                "id": novo_id,
            })

        if not vistas_geradas:
            tg.RollBack()
            forms.alert("Nenhuma vista criada.\n" + "\n".join(erros))
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
                            "NBR-6492" in nome_clean and "ESCALA" in nome_clean
                        ):
                            vp_type_id = vpt.Id
                            break
            except:
                continue

        with Transaction(doc, "Paginação de Pranchas — Isométricos") as t:
            t.Start()
            doc.Regenerate()

            mm  = 1.0 / 304.8
            margem_esq  = 25.0  * mm
            margem_dir  = 175.0 * mm
            margem_sup  = 10.0  * mm
            margem_inf  = 15.0  * mm
            espacamento = 15.0  * mm  # Espaçamento horizontal entre viewports
            DISTANCIA_TITULO_FT = -100.0 * mm  # Distância do título abaixo da vista (positivo = abaixo)
            MARGEM_LINHA = 20.0 * mm          # Margem vertical entre linhas (uniforme)

            def nova_prancha():
                sh = ViewSheet.Create(doc, tb_type_id)
                doc.Regenerate()
                tbs = (FilteredElementCollector(doc, sh.Id)
                       .OfCategory(BuiltInCategory.OST_TitleBlocks)
                       .WhereElementIsNotElementType().ToElements())
                tb_inst = tbs[0] if tbs else None

                for param_name, valor in dados_carimbo.items():
                    p = sh.LookupParameter(param_name)
                    if not p and tb_inst:
                        p = tb_inst.LookupParameter(param_name)
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

                if not tbs:
                    raise Exception("Prancha sem carimbo.")
                bb_sh = tb_inst.get_BoundingBox(sh)
                if not bb_sh:
                    raise Exception("BoundingBox da prancha é None.")
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
            vp_infos = []  # lista de dicts com vp, real_w, real_h por vista
            for info in vistas_geradas:
                v = doc.GetElement(info["id"])
                if not Viewport.CanAddViewToSheet(doc, sheet.Id, v.Id):
                    # Vista já está em outra prancha — cria nova prancha temporária
                    # apenas para medir; isso não acontece no fluxo normal, mas
                    # protege contra duplicatas acidentais.
                    erros.append("Vista '{}' não pode ser adicionada à prancha.".format(v.Name))
                    vp_infos.append(None)
                    continue

                vp = Viewport.Create(doc, sheet.Id, v.Id, XYZ(0, 0, 0))
                doc.Regenerate()

                if vp_type_id != ElementId.InvalidElementId:
                    vp.ChangeTypeId(vp_type_id)
                    doc.Regenerate()

                outline = vp.GetBoxOutline()
                real_w = outline.MaximumPoint.X - outline.MinimumPoint.X
                real_h = outline.MaximumPoint.Y - outline.MinimumPoint.Y

                vp_infos.append({
                    "vp":      vp,
                    "view_id": v.Id,   # guardado aqui pois Viewport não expõe .View no IronPython
                    "real_w":  real_w,
                    "real_h":  real_h,
                    "nome":    v.Name,
                })

            # ----------------------------------------------------------------
            # PASSO 2 — Agrupar em linhas usando largura REAL de cada viewport.
            # ----------------------------------------------------------------
            linhas = []
            linha_atual = []
            tmp_x = ux_min

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
                # Altura máxima real da linha
                rh_max = max(vi["real_h"] for vi in linha)

                # Verifica quebra de prancha (mesma constante MARGEM_LINHA)
                if cur_y - (rh_max + MARGEM_LINHA) < uy_min:
                    sheet, ux_min, ux_max, uy_min, uy_max = nova_prancha()
                    cur_y = uy_max

                    # Vistas já criadas na prancha anterior precisam ser
                    # recriadas na nova prancha — mover não é possível entre pranchas.
                    # Re-criar apenas as vistas desta linha e seguintes.
                    for vi in linha:
                        view_id = vi["view_id"]   # usa ID guardado — Viewport não expõe .View no IronPython
                        # Remove o viewport provisório da prancha anterior
                        try:
                            doc.Delete(vi["vp"].Id)
                            doc.Regenerate()
                        except:
                            pass
                        # Recria na nova prancha
                        new_vp = Viewport.Create(doc, sheet.Id, view_id, XYZ(0, 0, 0))
                        doc.Regenerate()
                        if vp_type_id != ElementId.InvalidElementId:
                            new_vp.ChangeTypeId(vp_type_id)
                            doc.Regenerate()
                        outline = new_vp.GetBoxOutline()
                        vi["vp"]     = new_vp
                        vi["real_w"] = outline.MaximumPoint.X - outline.MinimumPoint.X
                        vi["real_h"] = outline.MaximumPoint.Y - outline.MinimumPoint.Y

                    rh_max = max(vi["real_h"] for vi in linha)

                baseline_y = cur_y - rh_max
                cur_x = ux_min

                for vi in linha:
                    real_w = 0.0  # FIX: inicializado antes do try para evitar uso de valor anterior
                    try:
                        vp     = vi["vp"]
                        real_w = vi["real_w"]
                        real_h = vi["real_h"]

                        outline  = vp.GetBoxOutline()
                        shift_x  = cur_x - outline.MinimumPoint.X
                        shift_y  = baseline_y - outline.MinimumPoint.Y

                        DB.ElementTransformUtils.MoveElement(doc, vp.Id, XYZ(shift_x, shift_y, 0))
                        doc.Regenerate()

                        # Título posicionado abaixo da vista (DISTANCIA_TITULO_FT positivo = abaixo)
                        vp.LabelOffset = XYZ(0.0, -(real_h / 2.0) - DISTANCIA_TITULO_FT, 0)
                        doc.Regenerate()

                    except Exception as e:
                        erros.append("Erro ao posicionar vista '{}': {}".format(vi["nome"], str(e)))

                    cur_x += real_w + espacamento

                # Desce para a próxima linha com margem uniforme (MARGEM_LINHA)
                cur_y -= (rh_max + MARGEM_LINHA)

            t.Commit()
        tg.Assimilate()

    msg = "Concluído! {} isométrico(s) criado(s) e paginado(s).".format(len(vistas_geradas))
    if erros:
        msg += "\n\nAvisos ({}):\n{}".format(len(erros), "\n".join(erros))
    forms.alert(msg)


if __name__ == '__main__':
    executar_fluxo_isometrico()