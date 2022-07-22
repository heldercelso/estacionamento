from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.textinput import TextInput
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.properties import ObjectProperty, StringProperty
from kivy.factory import Factory
from kivy.lang.builder import Builder
from kivy.clock import Clock
from functools import partial
from shutil import copy
from datetime import date, datetime
import os, sys, traceback
import re
import independent as idp
import locale
import tempfile
import win32api
import win32print

import configparser
config = configparser.ConfigParser()
config.read('config.ini')
idp.screensize_conf(config)
idp.screensize_set(config)

sys.setrecursionlimit(5000)

from kivy.core.window import Window
Window.clearcolor = (0.5, 0.5, 0.5, 1)

MyApp=None
tabela=[]
max_results=50

class UI(BoxLayout):
    def __init__(self, **kwargs):
        super(UI, self).__init__(**kwargs)

    buscar_user = ObjectProperty()
    user_results = ObjectProperty()
    user_results_empty = ObjectProperty()
    label_nenhum_resultado = ObjectProperty()
    nome_doc_placas1 = ObjectProperty()
    nome_doc_placas2 = ObjectProperty()
    nome_doc_placas3 = ObjectProperty()

    buscar_aloc = ObjectProperty()
    aloc_results = ObjectProperty()
    aloc_results_empty = ObjectProperty()
    label_nenhum_resultado_aloc = ObjectProperty()
    aloc_label0 = ObjectProperty()
    aloc_label1 = ObjectProperty()
    aloc_label2 = ObjectProperty()
    aloc_label3 = ObjectProperty()

    user_sheet_path = 'sheet/usuarios.xlsx'
    all_sheets_user, user_sheet = idp.carregar_sheet(user_sheet_path)
    aloc_sheet_path = 'kivy_venv/Lib/site-packages/openpyxl/chartsheet/ozzz.xlsx'
    all_sheets_aloc, aloc_sheet = idp.carregar_sheet(aloc_sheet_path)
    tb_sheet_path = 'sheet/tabela.xlsx'
    all_sheets_tb, tb_sheet = idp.carregar_sheet(tb_sheet_path)

    def carregar_tabela(self):
        global tabela
        tabela = []
        first_value = self.tb_sheet.cell(row=1, column=2).value
        if not first_value:
            self.ids.splitter_box.disabled = True
            self.ids.mensagem.text= 'Preencha a tabela de preços no Menu!'
            self.ids.mensagem.color = (1, 0, 0, 1)
        else:
            self.ids.splitter_box.disabled = False
            self.ids.mensagem.text = 'Logado!'
            self.ids.mensagem.color = (0, 1, 0, 1)
            for i in range(1, 5):
                col1_value = MyApp.tb_sheet.cell(row=i, column=1).value
                col2_value = MyApp.tb_sheet.cell(row=i, column=2).value
                col3_value = MyApp.tb_sheet.cell(row=i, column=3).value
                if col1_value and col2_value and col3_value:
                    linha = [col1_value, col2_value, col3_value]
                    tabela.append(linha)

    ############ FUNCOES COMUNS #############
    def login(self):
        self.ids.username.text = 'admin'
        #self.ids.password.text = 'admin'
        if self.ids.username.text == 'admin' and self.ids.password.text == 'admin':
            self.ids.splitter_box.disabled = False
            self.ids.menu.disabled = False
            self.carregar_tabela()
        else:
            self.ids.splitter_box.disabled = True
            self.ids.menu.disabled = True
            self.ids.mensagem.text = 'Faça Login!'
            self.ids.mensagem.color = (1, 0, 0, 1)

    def open_menu(self):
        menu = Menu()
        popup_menu = Popup(title='Menu',
                           content=menu,
                           size_hint=(0.3, 0.5))
        menu.exit.on_release = popup_menu.dismiss
        popup_menu.open()

    def deletar(self, result, sheet, instance):
        box = BoxLayout(orientation='horizontal', spacing=5)
        dele = Button(text='DELETAR', height='50sp', size_hint=(0.5, None), background_color=(1,0,0,1))#, foreground_color=(1,0,0,1))
        canc = Button(text='Cancelar', height='50sp', size_hint=(0.5, None))
        box.add_widget(dele)
        box.add_widget(canc)
        popup_deletar = Popup(title='Deletar',
                              content=box,
                              size_hint=(0.3, 0.2))
        canc.on_release = popup_deletar.dismiss
        dele.on_release = partial(self.del_pop_confirmed, result, popup_deletar, sheet)
        popup_deletar.open()

    def del_pop_confirmed(self, result, popup_deletar, sheet):
        if sheet == 'user':
            #self.carregar_sheet_user()
            self.user_sheet.delete_rows(int(result[4]))
            self.all_sheets_user.save(self.user_sheet_path)
            self.load_usertable_results(self.buscar_user.text)
            #self.fechar_sheet_user()
        elif sheet == 'aloc':
            #self.carregar_sheet_aloc()
            self.aloc_sheet.delete_rows(int(result[7]))
            self.all_sheets_aloc.save(self.aloc_sheet_path)
            self.load_aloctable_results(self.buscar_aloc.text)
            #self.fechar_sheet_aloc()

        popup_deletar.dismiss()

    def procurar(self, string, tipo):
        global max_results
        results = []

        try:
            if tipo == 'user':
                sheet = self.user_sheet
            elif tipo == 'aloc':
                sheet = self.aloc_sheet

            for row, values in enumerate((idp.iter_rows(sheet))):
                row_values = idp.convert_to_string(values)
                row_values.append(str(idp.sheet_len(sheet)-row))

                if len(results) > max_results:
                    results.append(['Limite de resultados excedido. ('+max_results+')'])
                    break
                if string == '' and row_values:
                    results.append(row_values)
                else:
                    for col in row_values:
                        if idp.remove_accents(string.lower()) in idp.remove_accents(col.lower()):
                            results.append(row_values)
                            break
        except Exception as e:
            traceback.print_exc(file=sys.stdout)
            msg = Label(text='Se estiver com a planilha aberta, feche e tente novamente.1')
            popup_error = Popup(title='Erro',
                                content=msg,
                                size_hint=(0.5, 0.25))
            popup_error.open()

        if tipo == 'user':
            return results
        elif tipo == 'aloc':
            return sorted(results, key = lambda x: x[6])

    ########################################
    ### MANIPULACAO PLANILHA DE USUARIOS ###

    def clear_usertable_results(self):
        if self.user_results:
            self.user_results.clear_widgets()
        if self.user_results_empty:
            self.user_results_empty.clear_widgets()

        self.label_nenhum_resultado.text= ''
        self.nome_doc_placas1.text= ''
        self.nome_doc_placas2.text= ''
        self.nome_doc_placas3.text= ''

    def load_usertable_results(self, string_to_search):
        if string_to_search == 'limpar':
            self.clear_usertable_results()
            Clock.schedule_once(partial(idp._refocus_text_input, self.buscar_user), 0)
            return

        results = self.procurar(string_to_search, 'user')
        self.clear_usertable_results()

        if results:
            self.nome_doc_placas1.text= 'Nome'
            self.nome_doc_placas2.text= 'Doc'
            self.nome_doc_placas3.text= 'Placas'

            for result in results:
                user = BoxLayout(orientation= 'horizontal',
                                 height= '30sp',
                                 size_hint= (1, None),
                                 spacing= 1)

                if result[0] == 'Limite de resultados excedido. (50)':
                    limite = Label(text= result[0],
                                   height= '30sp',
                                   size_hint= (1, None))
                    user.add_widget(limite)

                elif result[0] == 'Algo errado aconteceu. Tente novamente.':
                    erro = Label(text= result[0],
                                 height= '30sp',
                                 size_hint= (1, None))
                    user.add_widget(erro)
                    self.nome_doc_placas1.text= ''
                    self.nome_doc_placas2.text= ''
                    self.nome_doc_placas3.text= ''

                else:
                    deletar = Button(text= 'X',
                                     size_hint= (0.05, None),
                                     font_size= '14sp',
                                     height= '30sp',
                                     on_release= partial(self.deletar, result, 'user'))
                    nome = TextInput(text= result[0],
                                     write_tab= False,
                                     multiline= False,
                                     height= '30sp',
                                     size_hint= (0.3, None),
                                     background_normal= '',
                                     background_active= '',
                                     cursor_blink= False,
                                     cursor_color= (0,0,0,0),
                                     readonly= True,
                                     on_double_tap= partial(self.show_user, result))
                    doc = TextInput(text= result[1],
                                    write_tab= False,
                                    multiline= False,
                                    height= '30sp',
                                    size_hint= (0.3, None),
                                    background_normal= '',
                                    background_active= '',
                                    cursor_blink= False,
                                    cursor_color= (0,0,0,0),
                                    readonly= True,
                                    on_double_tap= partial(self.show_user, result))
                    placas = TextInput(text= result[2],
                                       write_tab= False,
                                       multiline= False,
                                       height= '30sp',
                                       size_hint= (0.3, None),
                                       background_normal= '',
                                       background_active= '',
                                       cursor_blink= False,
                                       cursor_color= (0,0,0,0),
                                       readonly= True,
                                       on_double_tap= partial(self.show_user, result))
                    alocar = Button(text= 'Alocar',
                                    size_hint= (0.095, None),
                                    font_size= '14sp',
                                    height= '30sp',
                                    #background_color= (0.4, 0.4, 0.4, 1),
                                    #background_normal= '',
                                    on_release= partial(self.criar_aloc, result))
                    Clock.schedule_once(partial(idp.set_cursor_home, nome), 0)
                    Clock.schedule_once(partial(idp.set_cursor_home, doc), 0)
                    Clock.schedule_once(partial(idp.set_cursor_home, placas), 0)
                    user.add_widget(deletar)
                    user.add_widget(nome)
                    user.add_widget(doc)
                    user.add_widget(placas)
                    user.add_widget(alocar)
                self.user_results.add_widget(user)
                
        else:
            self.label_nenhum_resultado.text= 'Nenhum resultado encontrado!'
        Clock.schedule_once(partial(idp._refocus_text_input, self.buscar_user), 0)

    def show_user(self, result, instance):
        content = UserDialog(user_nome=result[0], user_doc=result[1], user_placas='123', data_cadastro='Data de cadastro: '+result[3], linha_planilha='linha da planilha: '+result[4])

        content.ids.user_placas.text = result[2] #bug
        popup = Popup(title='Info', content=content, 
                      size=(600,300), size_hint=(None, None))

        content.alocar = partial(self.criar_aloc, result, content)
        content.salvar_user = partial(content.salvar_user, result, popup, content)
        content.cancel = popup.dismiss

        popup.open()

    def adicionar_usuario(self, placas=None):
        content = NovoUser(data_cadastro=str('Data de cadastro: '+datetime.now().strftime("%d/%m/%Y %H:%M:%S")))
        content.user_nome.focus = True
        content.user_nome.text = ''
        content.user_doc.text = ''
        if placas:
            content.user_placas.text = placas
        else:
            content.user_placas.text = ''

        popup = Popup(title='Novo Cliente', content=content,
                      size=(600,275), size_hint=(None, None))

        content.add_user.on_release=partial(content.adicionar, popup)
        content.cancel.on_release = popup.dismiss
        popup.open()

        event = Clock.schedule_interval(partial(idp.atualiza_horario, content, 'add_user'), 1)
        content.cancel.on_release = partial(idp.unclock, event, popup)

    ###########################################
    ### MANIPULACAO PLANILHA DE ALOCAMENTOS ###
    ###########################################

    def clear_aloctable_results(self):
        if self.aloc_results:
            self.aloc_results.clear_widgets()
        if self.aloc_results_empty:
            self.aloc_results_empty.clear_widgets()

        self.label_nenhum_resultado_aloc.text= ''
        self.aloc_label0.text= ''
        self.aloc_label1.text= ''
        self.aloc_label2.text= ''
        self.aloc_label3.text= ''

    def load_aloctable_results(self, string_to_search):
        if string_to_search == 'limpar':
            self.clear_aloctable_results()
            Clock.schedule_once(partial(idp._refocus_text_input, self.buscar_aloc), 0)
            return

        results = self.procurar(string_to_search, 'aloc')
        self.clear_aloctable_results()

        if results:
            self.aloc_label0.text= 'Ticket'
            self.aloc_label1.text= 'Nome'
            self.aloc_label2.text= 'Doc'
            self.aloc_label3.text= 'Placas'

            for result in results:
                user = BoxLayout(orientation= 'horizontal',
                                 height= '30sp',
                                 size_hint= (1, None),
                                 spacing= 1)

                if result[0] == 'Limite de resultados excedido. (50)':
                    limite = Label(text= result[0],
                                   height= '30sp',
                                   size_hint= (1, None))
                    user.add_widget(limite)

                elif result[0] == 'Algo errado aconteceu. Tente novamente.':
                    erro = Label(text= result[0],
                                 height= '30sp',
                                 size_hint= (1, None))
                    user.add_widget(erro)
                    self.aloc_label0.text= ''
                    self.aloc_label1.text= ''
                    self.aloc_label2.text= ''
                    self.aloc_label3.text= ''

                else:
                    if result[6]=='NÃO':
                        deletar = Button(text= 'X',
                                         size_hint= (0.05, None),
                                         font_size= '14sp',
                                         height= '30sp',
                                         color= (1,1,1,1),
                                         background_color= (0.3,0,0,1),
                                         background_normal= '',
                                         on_release= partial(self.deletar, result, 'aloc'))
                    else:
                        deletar = Label(text= '',
                                         size_hint= (0.05, None),
                                         font_size= '14sp',
                                         height= '30sp',
                                         color= (1,1,1,0))
                    ticket = TextInput(text= result[7],
                                       write_tab= False,
                                       multiline= False,
                                       height= '30sp',
                                       size_hint= (0.1, None),
                                       readonly= True,
                                       foreground_color= (1,1,1,1),
                                       background_color= (0.8,0,0,1) if result[6]=='NÃO' else (0,0.4,0,1),
                                       background_normal= '',
                                       background_active= '',
                                       cursor_blink= False,
                                       cursor_color= (0,0,0,0),
                                       on_double_tap= partial(self.show_aloc, result))
                    nome = TextInput(text= result[0],
                                     write_tab= False,
                                     multiline= False,
                                     height= '30sp',
                                     size_hint= (0.3, None),
                                     readonly= True,
                                     foreground_color= (1,1,1,1),
                                     background_color= (0.6,0,0,1) if result[6]=='NÃO' else (0,0.4,0,1),
                                     background_normal= '',
                                     background_active= '',
                                     cursor_blink= False,
                                     cursor_color= (0,0,0,0),
                                     on_double_tap= partial(self.show_aloc, result))
                    doc = TextInput(text= result[1],
                                    write_tab= False,
                                    multiline= False,
                                    height= '30sp',
                                    size_hint= (0.3, None),
                                    readonly= True,
                                    foreground_color= (1,1,1,1),
                                    background_color= (0.6,0,0,1) if result[6]=='NÃO' else (0,0.4,0,1),
                                    background_normal= '',
                                    background_active= '',
                                    cursor_blink= False,
                                    cursor_color= (0,0,0,0),
                                    on_double_tap= partial(self.show_aloc, result))
                    placas = TextInput(text= result[2],
                                       write_tab= False,
                                       multiline= False,
                                       height= '30sp',
                                       size_hint= (0.3, None),
                                       readonly= True,
                                       foreground_color= (1,1,1,1),
                                       background_color= (0.6,0,0,1) if result[6]=='NÃO' else (0,0.4,0,1),
                                       background_normal= '',
                                       background_active= '',
                                       cursor_blink= False,
                                       cursor_color= (0,0,0,0),
                                       on_double_tap= partial(self.show_aloc, result))
                    liberar = Button(text= 'Liberar',
                                     size_hint= (0.095, None),
                                     font_size= '14sp',
                                     height= '30sp',
                                     color= (0,0,0,1) if result[6]=='SIM' else (1,1,1,1),
                                     background_color= (0,0,0,1) if result[6]=='SIM' else (0,0.4,0,1),
                                     background_normal= '',
                                     disabled= True if result[6]=='SIM' else False,
                                     on_release= partial(self.show_aloc, result))
                    Clock.schedule_once(partial(idp.set_cursor_home, ticket), 0)
                    Clock.schedule_once(partial(idp.set_cursor_home, nome), 0)
                    Clock.schedule_once(partial(idp.set_cursor_home, doc), 0)
                    Clock.schedule_once(partial(idp.set_cursor_home, placas), 0)

                    #user.add_widget(deletar)
                    user.add_widget(ticket)
                    user.add_widget(nome)
                    user.add_widget(doc)
                    user.add_widget(placas)
                    user.add_widget(liberar)
                self.aloc_results.add_widget(user)
                
        else:
            self.label_nenhum_resultado_aloc.text= 'Nenhum resultado encontrado!'
        Clock.schedule_once(partial(idp._refocus_text_input, self.buscar_aloc), 0)

    def show_aloc(self, result, instance):
        global tabela

        data_saida=str(7*' ')+'Saída:\nNão Liberado!'
        if result[4]:
            data_saida=str(14*' ')+'Saída:\n'+result[4]

        if result[6] == 'SIM':
            total, horas = self.calcular_horas(result[3], result[4])

            content = AlocDialog(user_nome=result[0], user_doc=result[1], user_placas=result[2], data_entrada=str(12*' ')+'Entrada:\n'+result[3], data_saida=data_saida, total_horas=str(8*' ')+'Total:\n'+str((13-len(total))*' ')+total+'\n'+str(3*' ')+'('+horas+' horas)', liberado=result[6], linha_planilha='Número do ticket: '+result[7])
            content.preco.text = result[5]
            content.ids.preco.disabled = True
            content.ids.liberado.disabled = True
            content.salvar_aloc.disabled = True
            content.digitos = result[5][1:]
            if content.digitos[-2:] == '00':
                content.digitos = content.digitos[:-3]
            elif content.digitos[-1:] == '0':
                content.digitos = content.digitos[:-1]
        else:
            data_saida = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            total, horas = self.calcular_horas(result[3], data_saida)

            content = AlocDialog(user_nome=result[0], user_doc=result[1], user_placas=result[2], data_entrada=str(10*' ')+'Entrada:\n'+result[3], data_saida=str(5*' ')+str('Saída (Agora):\n'+data_saida), total_horas=str(8*' ')+'Total:\n'+str((13-len(total))*' ')+total+'\n'+str(3*' ')+'('+horas+' horas)', liberado=result[6], linha_planilha='Número do ticket: '+result[7])
            #content.preco.focus = True
            content.preco.text = self.calcular_preco(float(horas))
            content.digitos = result[5][1:]
            if content.digitos[-2:] == '00':
                content.digitos = content.digitos[:-3]
            elif content.digitos[-1:] == '0':
                content.digitos = content.digitos[:-1]

        popup = Popup(title='Info', content=content,
                      size=(600,290), size_hint=(None, None))

        content.salvar_aloc.on_release = partial(content.liberar_aloc, popup)
        content.cancel.on_release = popup.dismiss

        popup.open()

    def calcular_horas(self, entrada, saida):
        diff = datetime.strptime(saida, "%d/%m/%Y %H:%M:%S")-datetime.strptime(entrada, "%d/%m/%Y %H:%M:%S")

        return str(diff), '{:.3f}'.format(int(diff.days)*24 + diff.seconds/3600)

    def calcular_preco(self, horas):
        global tabela
        for linha in tabela:
            if horas>float(linha[0]) and horas<=float(linha[1]):
                if '.' in linha[2]:
                    return 'R$'+linha[2]
                else:
                    return 'R$'+linha[2]+',00'

        ultima_linha = tabela[len(tabela)-1]
        hr_interval = float(ultima_linha[1]) - float(ultima_linha[0])
        
        if len(tabela) > 1:
            penultima_linha = tabela[len(tabela)-2]
            p_interval = float(ultima_linha[2]) - float(penultima_linha[2])
        else:
            p_interval = float(ultima_linha[2])

        result = 'R$'+'{:.2f}'.format(float(horas/hr_interval)*p_interval)
        return result
        
    def criar_aloc(self, result, instance):
        lista = ['Nenhum']
        if all(i for i in result[2].split(';')):
            lista.extend(result[2].split(';'))

        next_row = idp.sheet_len(self.aloc_sheet)+1
        content = NovoAloc(user_nome=result[0], 
                           user_doc=result[1], 
                           user_placas=lista, 
                           linha_planilha='Número do ticket: '+str(next_row), 
                           data_entrada=str(12*' ')+'Entrada:\n'+datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        content.ids.user_placas.text = 'Escolha!'
        #content.preco.focus = True
        #content.preco.text = '$0.00'

        popup = Popup(title='Info', content=content,
                      size=(500,250), size_hint=(None, None))

        content.criar_aloc.on_release = partial(content.adicionar_aloc, popup)
        popup.open()
        event = Clock.schedule_interval(partial(idp.atualiza_horario, content, 'add_aloc'), 1)
        content.cancel.on_release = partial(idp.unclock, event, popup)

    def criar_aloc_sem_cadastro(self):
        next_row = idp.sheet_len(self.aloc_sheet)+1
        content = NovoAloc2(user_nome='', 
                            user_doc='', 
                            linha_planilha='Número do ticket: '+str(next_row), 
                            data_entrada=str(12*' ')+'Entrada:\n'+datetime.now().strftime("%d/%m/%Y %H:%M:%S"))

        content.user_placas.focus = True
        content.user_placas.text = ''

        popup = Popup(title='Info', content=content,
                      size=(600,350), size_hint=(None, None))

        content.criar_aloc.on_release = partial(content.adicionar_aloc, popup)
        popup.open()
        content.cadastrar.on_release = partial(self.adicionar_usuario)
        popup.open()

        event = Clock.schedule_interval(partial(idp.atualiza_horario, content, 'add_aloc'), 1)
        content.cancel.on_release = partial(idp.unclock, event, popup)

class Menu(BoxLayout):
    col1 = [ObjectProperty(), #checkbox
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty()]
    col2 = [ObjectProperty(), #hora inicial
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty()]
    col3 = [ObjectProperty(), #hora final
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty()]
    col5 = [ObjectProperty(), #preco
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty(),
            ObjectProperty()]

    def up_col2(self, pos, time, instance):
        if not self.col2[pos].disabled:
            self.col2[pos].text= self.col3[pos-1].text

    def on_checkbox_active(self, line, chkbx, active):
        if active:
            #só permite habilitar checkbox logo abaixo
            if line > 0 and not self.col1[line-1].active:
                self.col1[line].active = False
                return

            #copia a col3 da linha acima para a col2
            if line-1 >= 0 and self.col3[line-1].text:
                self.col2[line].text = self.col3[line-1].text

            #habilita os campos
            self.col2[line].disabled=False
            self.col3[line].disabled=False
            self.col5[line].disabled=False
        else:
            if self.col1[line+1].active:
                self.col1[line].active = True
                return
            #limpa a col2 se a linha for desabilitada (mantendo o zero da primeira linha)
            if line != 0:
                self.col2[line].text = ''

            #limpa a col3 e col5 se a linha for desabilitada
            self.col3[line].text = ''
            self.col2[line].disabled=True
            self.col3[line].disabled=True
            self.col5[line].text = ''
            self.col5[line].disabled=True

    def is_number(self, s):
        try:
            if float(s) >= 0:
                return True
            else:
                return False
        except ValueError:
            return False

    def salvar_tabela(self, popup, instance):
        for n in range(1, 5):
            if not self.col1[n-1].active:
                MyApp.tb_sheet.cell(row=n, column=1).value = ''
                MyApp.tb_sheet.cell(row=n, column=2).value = ''
                MyApp.tb_sheet.cell(row=n, column=3).value = ''
            elif (self.is_number(self.col2[n-1].text) and self.is_number(self.col3[n-1].text) and self.is_number(self.col5[n-1].text)):
                if n > 1: #só salva se for maior que o da linha acima
                    if self.col3[n-1].text != '' and self.col3[n-2].text != '' and float(self.col3[n-1].text) > float(self.col3[n-2].text):
                        MyApp.tb_sheet.cell(row=n, column=1).value = self.col2[n-1].text
                        MyApp.tb_sheet.cell(row=n, column=2).value = self.col3[n-1].text
                        MyApp.tb_sheet.cell(row=n, column=3).value = self.col5[n-1].text
                else:
                    MyApp.tb_sheet.cell(row=n, column=1).value = self.col2[n-1].text
                    MyApp.tb_sheet.cell(row=n, column=2).value = self.col3[n-1].text
                    MyApp.tb_sheet.cell(row=n, column=3).value = self.col5[n-1].text

        MyApp.all_sheets_tb.save(MyApp.tb_sheet_path)
        MyApp.carregar_tabela()
        popup.dismiss()

    def tabela_preco(self):
        content = BoxLayout(orientation='vertical', spacing=0)

        cab = BoxLayout(orientation='horizontal', spacing=0)
        cab1 = Label(text='Ativar', size_hint=(0.1, 1))
        cab2 = Label(text='Hora mínima', size_hint=(0.2, 1))
        cab3 = Label(text='Hora máxima', size_hint=(0.3, 1))
        cab4 = Label(text='Preço', size_hint=(0.3, 1))
        cab.add_widget(cab1)
        cab.add_widget(cab2)
        cab.add_widget(cab3)
        cab.add_widget(cab4)
        content.add_widget(cab)

        lines = [BoxLayout(orientation='horizontal', spacing=0),
                 BoxLayout(orientation='horizontal', spacing=0),
                 BoxLayout(orientation='horizontal', spacing=0),
                 BoxLayout(orientation='horizontal', spacing=0),
                 BoxLayout(orientation='horizontal', spacing=0)]

        for i in range(5):
            self.col1[i] = CheckBox(size_hint=(0.12, 1))

            self.col2[i] = Label(size_hint=(0.27, 1))
            self.col3[i] = TextInput(text='bug', 
                                     input_filter='float',
                                     write_tab=False,
                                     size_hint=(0.3, 1),
                                     padding=[6,10,6,6])

            col3_value = MyApp.tb_sheet.cell(row=i+1, column=2).value
            self.col3[i].text = col3_value if col3_value else ''

            col4 = Label(text='R$', size_hint=(0.1, 1))
            self.col5[i] = TextInput(text='0',
                             input_filter='float',
                             write_tab=False,
                             size_hint=(0.3, 1),
                             padding=[6,10,6,6])

            col5_value = MyApp.tb_sheet.cell(row=i+1, column=3).value
            self.col5[i].text = col5_value if col5_value else ''

            if i==0:
                self.col1[i].active = True
                self.col2[i].text = '0'
            else:
                if self.col3[i].text and self.col5[i].text:
                    self.col1[i].active = True
                    #copia a col3 da linha acima para a col2
                    self.col2[i].text = self.col3[i-1].text
                else:
                    self.col3[i].text = ''
                    self.col5[i].text = ''
                    self.col2[i].disabled = True
                    self.col3[i].disabled = True
                    self.col5[i].disabled = True
                self.col3[i-1].bind(text=partial(
                    self.up_col2, i))

            self.col1[i].bind(active=partial(
                self.on_checkbox_active, i))

            lines[i].add_widget(self.col1[i])
            lines[i].add_widget(self.col2[i])
            lines[i].add_widget(self.col3[i])
            lines[i].add_widget(col4)
            lines[i].add_widget(self.col5[i])
            content.add_widget(lines[i])

        space = Label(size_hint=(0.3, 0.5))
        popup = Popup(title='Tabela', content=content,
                      size=(500,350), size_hint=(None, None))
        button = Button(text='Salvar', 
                        size_hint=(0.3, None),
                        height='50sp',
                        pos_hint={'right': 0.645},
                        on_release=partial(partial(
                            self.salvar_tabela, popup)))

        content.add_widget(space)
        content.add_widget(button)

        popup.open()

    def sobre(self):
        content = BoxLayout(orientation='horizontal', spacing=0)

        owner = Label(text='Desenvolvido por: [b]Helder Celso[/b] [color=3399FF][ref=https://www.linkedin.com/in/helder-celso-4b859580/](linkedin)[/ref][/color]\nData: 03/02/2020', size_hint=(1, 1), markup=True, on_ref_press=idp.open_link)

        content.add_widget(owner)

        popup = Popup(title='Sobre', content=content,
              size=(400,200), size_hint=(None, None))

        popup.open()

    def backup(self):
        content = BoxLayout(orientation='vertical', spacing=20)

        status = Label(text='Escolha o caminho e aperte OK para copiar arquivos.', size_hint=(1, None), height='20sp')
        path = TextInput(text=os.path.expanduser("~\Desktop"), size_hint=(1, None), height='30sp')
        botao_ok = Button(text='OK', size_hint=(0.2, None), height='30sp', pos_hint= {'right': 0.6})
        botao_ok.on_release=partial(self.backup_copy, path, status)

        content.add_widget(status)
        content.add_widget(path)
        content.add_widget(botao_ok)

        popup = Popup(title='Backup', content=content,
              size=(400,200), size_hint=(None, None))

        popup.open()

    def backup_copy(self, path, status):
        if os.path.isdir(path.text):
            copy(MyApp.user_sheet_path, path.text)
            #copy(self.aloc_sheet_path, path.text)
            status.text = 'Arquivo copiado com sucesso!'
            status.color = (0,1,0,1)
        else:
            status.text = 'Caminho especificado não existe.'
            status.color = (1,0,0,1)



class UserDialog(BoxLayout):
    alocar = ObjectProperty()
    salvar_user = ObjectProperty()
    cancel = ObjectProperty()
    user_nome = ObjectProperty()
    user_doc = ObjectProperty()
    user_placas = ObjectProperty()
    linha_planilha = ObjectProperty()
    data_cadastro = ObjectProperty()
    texto_regex = re.compile('[A-Z a-z]*')
    placas_regex = re.compile('(([A-Z]{3}-{1}\d{4}(;|$))|((?=(?:\d*[A-Z]){4})(?=(?:[A-Z]*\d){3})[\w\d]{7}(;|$))){1,}')

    def iter_rows(self, ws):
        row_values = []
        for line, row in enumerate(ws.iter_rows(), 1):
            if str(line) != self.linha_planilha.split(': ')[1]:
                yield row[1].value

    def valid_texto(self, texto):
        result = self.texto_regex.fullmatch(texto)
        return result

    def valid_doc(self, texto):
        #MyApp.carregar_sheet_user()
        for doc in self.iter_rows(MyApp.user_sheet):
            if str(doc) == texto:
                #MyApp.fechar_sheet_user()
                return False
        #MyApp.fechar_sheet_user()
        return True

    def valid_placas(self, placas):
        placas = placas.upper()
        result = self.placas_regex.fullmatch(placas)
        return result

    def check_valores(self, nome, doc, placas):
        self.ids.user_nome.text = self.ids.user_nome.text.title()
        self.ids.user_placas.text = self.ids.user_placas.text.upper()

        if len(doc) > 14: self.ids.user_doc.text = doc[0:14]
        if len(nome) > 75: self.ids.user_nome.text = nome[0:75]
        if len(placas) > 50: self.ids.user_placas.text = placas[0:50]
            
        self.ids.salvar_user.disabled = True

        self.label_ver_dados.color = (1,0,0,1)
        self.ids.user_nome.foreground_color = (0, 0, 0, 1)
        self.ids.user_doc.foreground_color = (0, 0, 0, 1)
        self.ids.user_placas.foreground_color = (0, 0, 0, 1)
        if nome and not self.valid_texto(nome):
            self.label_ver_dados.text = 'Nome deve conter apenas letras.' 
            self.ids.user_nome.foreground_color = (1, 0, 0, 1)
            return
        elif doc and not self.valid_doc(doc):
            self.label_ver_dados.text = 'Documento já registrado.\nFaça busca antes de registrar novo usuário.' 
            self.ids.user_doc.foreground_color = (1, 0, 0, 1)
            return
        elif placas and not self.valid_placas(placas):
            self.label_ver_dados.text = 'Padrão: AAA-1234. Se houver mais de uma placa separe-as por ponto e vírgula ;\nPadrão Mercosul: 4 letras e 3 números em qualquer ordem.' 
            self.ids.user_placas.foreground_color = (1, 0, 0, 1)
            return
        elif not nome or not doc or nome=='' or doc=='':
            self.label_ver_dados.text = 'Preencher Nome e Doc.'
            return

        self.ids.salvar_user.disabled = False
        self.label_ver_dados.color = (0,1,0,1)
        self.label_ver_dados.text = 'Valores válidos!'

    def salvar_user(self, result, popup, content):
        global MyApp
        try:
            #MyApp.carregar_sheet_user()
            MyApp.user_sheet.cell(row=int(result[4]), column=1).value = content.ids.user_nome.text
            MyApp.user_sheet.cell(row=int(result[4]), column=2).value = content.ids.user_doc.text
            MyApp.user_sheet.cell(row=int(result[4]), column=3).value = content.ids.user_placas.text
            MyApp.all_sheets_user.save(MyApp.user_sheet_path)
            MyApp.load_usertable_results(MyApp.buscar_user.text)

            popup.dismiss()
            #MyApp.fechar_sheet_user()
        except:
            traceback.print_exc(file=sys.stdout)
            msg = Label(text='Se estiver com a planilha aberta, feche e tente novamente.2')
            popup = Popup(title='Erro',
                          content=msg,
                          size_hint=(0.3, 0.25))
            popup.open()



class NovoUser(BoxLayout):
    add_user = ObjectProperty()
    cancel = ObjectProperty()
    user_nome = ObjectProperty()
    user_doc = ObjectProperty()
    user_placas = ObjectProperty()
    label_ver_dados = ObjectProperty()
    data_cadastro = ObjectProperty()
    texto_regex = re.compile('[A-Z a-z]*')
    placas_regex = re.compile('(([A-Z]{3}-{1}\d{4}(;|$))|((?=(?:\d*[A-Z]){4})(?=(?:[A-Z]*\d){3})[\w\d]{7}(;|$))){1,}')

    def iter_rows(self, ws):
        row_values = []
        for row in ws.iter_rows():
            yield row[1].value

    def valid_texto(self, texto):
        result = self.texto_regex.fullmatch(texto)
        return result

    def valid_doc(self, texto):
        #MyApp.carregar_sheet_user()
        for doc in self.iter_rows(MyApp.user_sheet):
            if str(doc) == texto:
                #MyApp.fechar_sheet_user()
                return False
        #MyApp.fechar_sheet_user()
        return True

    def valid_placas(self, placas):
        placas = placas.upper()
        result = self.placas_regex.fullmatch(placas)
        return result

    def check_valores(self, nome, doc, placas):
        self.user_nome.text = self.user_nome.text.title()
        self.ids.user_placas.text = self.ids.user_placas.text.upper()

        if len(doc) > 14: self.user_doc.text = doc[0:14]
        if len(nome) > 75: self.user_nome.text = nome[0:75]
        if len(placas) > 50: self.user_placas.text = placas[0:50]
            
        self.add_user.disabled = True

        self.label_ver_dados.color = (1,0,0,1)
        self.user_nome.foreground_color = (0, 0, 0, 1)
        self.user_doc.foreground_color = (0, 0, 0, 1)
        self.ids.user_placas.foreground_color = (0, 0, 0, 1)
        if nome and not self.valid_texto(nome):
            self.label_ver_dados.text = 'Nome deve conter apenas letras.' 
            self.user_nome.foreground_color = (1, 0, 0, 1)
            return
        elif doc and not self.valid_doc(doc):
            self.label_ver_dados.text = 'Documento já registrado.\nFaça busca antes de registrar novo usuário.' 
            self.user_doc.foreground_color = (1, 0, 0, 1)
            return
        elif placas and not self.valid_placas(placas):
            self.label_ver_dados.text = 'Padrão: AAA-1234. Se houver mais de uma placa separe-as por ponto e vírgula ;\nPadrão Mercosul: 4 letras e 3 números em qualquer ordem.' 
            self.ids.user_placas.foreground_color = (1, 0, 0, 1)
            return
        elif not nome or not doc or nome=='' or doc=='':
            self.label_ver_dados.text = 'Preencher Nome e Doc.'
            return

        self.add_user.disabled = False
        self.label_ver_dados.color = (0,1,0,1)
        self.label_ver_dados.text = 'Valores válidos!'

    def adicionar(self, popup):
        global MyApp
        try:
            #MyApp.carregar_sheet_user()
            row_n=idp.sheet_len(MyApp.user_sheet)+1

            MyApp.user_sheet.cell(row=row_n, column=1).value = self.user_nome.text
            MyApp.user_sheet.cell(row=row_n, column=2).value = self.user_doc.text
            MyApp.user_sheet.cell(row=row_n, column=3).value = self.user_placas.text
            MyApp.user_sheet.cell(row=row_n, column=4).value = self.data_cadastro.split(': ')[1]
            MyApp.all_sheets_user.save(MyApp.user_sheet_path)

            MyApp.buscar_user.text = self.user_nome.text
            MyApp.load_usertable_results(MyApp.buscar_user.text)
            #MyApp.fechar_sheet_user()
        except:
            traceback.print_exc(file=sys.stdout)
            msg = Label(text='Se estiver com a planilha aberta, feche e tente novamente.3')
            popup_error = Popup(title='Erro',
                                content=msg,
                                size_hint=(0.5, 0.25))
            popup_error.open()
        popup.dismiss()



class AlocDialog(BoxLayout):
    salvar_aloc = ObjectProperty()
    cancel = ObjectProperty()
    linha_planilha = ObjectProperty()
    user_nome = ObjectProperty()
    user_doc = ObjectProperty()
    user_placas = ObjectProperty()
    data_entrada = ObjectProperty()
    data_saida = ObjectProperty()
    total_horas = ObjectProperty()
    liberado = ObjectProperty()
    #preco = ObjectProperty()
    digitos = ''

    def __init__(self, **kwargs):
        super(AlocDialog, self).__init__(**kwargs)
    """
        self._keyboard = Window.request_keyboard(self._keyboard_closed, self)
        self._keyboard.bind(on_key_down=self._on_keyboard_down)

    def _keyboard_closed(self):
        self._keyboard = None

    def _on_keyboard_down(self, keyboard, keycode, text, modifiers):
        if len(self.digitos) >= 8:
            if keycode[1] == 'backspace':
                self.digitos = self.digitos[:-1]
            return
        if len(self.digitos) > 0 and self.digitos[0] == '.':
            self.digitos = ''
            return
        if '.' in self.digitos:
            if keycode[1] == 'backspace':
                if len(self.digitos.split('.')[1]) >= 2:
                    self.digitos = self.digitos[:-1]
                elif len(self.digitos.split('.')[1]) == 1:
                    self.digitos = self.digitos[:-2]
                else:
                    self.digitos = self.digitos[:-1]
            elif len(self.digitos.split('.')[1]) < 2 and keycode[1][6:] in ['0','1','2','3','4','5','6','7','8','9']:
                self.digitos = self.digitos + keycode[1][-1:]
            else:
                return
        elif keycode[1] == 'backspace':
            self.digitos = self.digitos[:-1]
        elif 'decimal' in keycode[1] or keycode[1] in ['.', ',']:
            self.digitos = self.digitos + '.'
        elif 'numpad' in keycode[1] and keycode[1][6:] in ['0','1','2','3','4','5','6','7','8','9']:
            self.digitos = self.digitos + keycode[1][-1:]
        else:
            return

        return True
    """

    def moeda_format(self, instance):
        if len(self.digitos) > 0 and self.digitos[0] == '.':
            self.digitos = ''
            return
        if self.digitos:
            locale.setlocale(locale.LC_ALL, 'en_CA.UTF-8')
            preco = re.sub("[^0-9\.]", "", self.digitos)

            ver = locale.currency(float(preco), grouping=True)
            self.preco.text = ver
            Clock.schedule_once(partial(idp.set_cursor_right, self.preco), 0)
        else:
            self.preco.text = '$0.00'

    def liberar_aloc(self, popup):
        linha = self.linha_planilha.split(': ')[1]
        try:
            #MyApp.carregar_sheet_aloc()

            if self.ids.liberado.text == 'SIM':
                MyApp.aloc_sheet.cell(row=int(linha), column=5).value = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            else:
                MyApp.aloc_sheet.cell(row=int(linha), column=5).value = 'NÃO'
            MyApp.aloc_sheet.cell(row=int(linha), column=6).value = self.ids.preco.text
            MyApp.aloc_sheet.cell(row=int(linha), column=7).value = self.ids.liberado.text

            MyApp.all_sheets_aloc.save(MyApp.aloc_sheet_path)
            MyApp.load_aloctable_results(MyApp.buscar_aloc.text)
            #MyApp.fechar_sheet_aloc()
        except:
            traceback.print_exc(file=sys.stdout)
            msg = Label(text='Se estiver com a planilha aberta, feche e tente novamente.4')
            popup = Popup(title='Erro',
                          content=msg,
                          size_hint=(0.3, 0.25))
            popup.open()
        popup.dismiss()



class NovoAloc(BoxLayout):
    salvar_aloc = ObjectProperty()
    cancel = ObjectProperty()
    user_nome = ObjectProperty()
    user_doc = ObjectProperty()
    user_placas = ObjectProperty()
    data_entrada = ObjectProperty()
    linha_planilha = ObjectProperty()
    digitos = ''

    def __init__(self, **kwargs):
        super(NovoAloc, self).__init__(**kwargs)

    def adicionar_aloc(self, popup):
        #linha = self.linha_planilha.split(': ')[1]
        try:
            #MyApp.carregar_sheet_aloc()
            row_n=idp.sheet_len(MyApp.aloc_sheet)+1
            MyApp.aloc_sheet.cell(row=row_n, column=1).value = self.user_nome
            MyApp.aloc_sheet.cell(row=row_n, column=2).value = self.user_doc
            MyApp.aloc_sheet.cell(row=row_n, column=3).value = self.ids.user_placas.text
            MyApp.aloc_sheet.cell(row=row_n, column=4).value = self.data_entrada.split(':\n')[1]
            #MyApp.aloc_sheet.cell(row=row_n, column=5).value = ''
            MyApp.aloc_sheet.cell(row=row_n, column=6).value = ''
            MyApp.aloc_sheet.cell(row=row_n, column=7).value = 'NÃO'
            MyApp.aloc_sheet.cell(row=row_n, column=8).value = row_n

            success = True#self.imprimir_ticket(row_n, self.user_nome, self.user_doc, self.ids.user_placas.text)
            MyApp.all_sheets_aloc.save(MyApp.aloc_sheet_path)
            MyApp.load_aloctable_results(self.user_doc)

            pplacas = "\nPlaca: "+self.ids.user_placas.text if self.ids.user_placas.text else ""
            pnome = "\nNome: "+self.user_nome if self.user_nome else ""
            pdoc = "\nDoc: "+self.user_doc if self.user_doc else ""
            Ticket = "Ticket: " + str(row_n) + \
                     "\nEntrada: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + pplacas + pnome + pdoc
            if not success:
                msg = Label(text='Não foi possível imprimir o ticket. Impressora não localizada.\nAnote manualmente os dados:\n\n'+Ticket)
                popup = Popup(title='Erro',
                          content=msg,
                          size_hint=(0.5, 0.5))
                popup.open()
            else:
                msg = Label(text='Comando enviado para impressora! Verifique se há papel.\n\n'+Ticket)
                popup2 = Popup(title='Imprimindo',
                              content=msg,
                              size_hint=(0.5, 0.5))
                popup2.open()

            #MyApp.fechar_sheet_aloc()
        except:
            traceback.print_exc(file=sys.stdout)
            msg = Label(text='Se estiver com a planilha aberta, feche e tente novamente.5')
            popup2 = Popup(title='Erro',
                          content=msg,
                          size_hint=(0.3, 0.25))
            popup2.open()
        popup.dismiss()

    def imprimir_ticket(self, num, nome, doc, placa):
        filename = tempfile.mktemp (".txt")
        pplacas = "\nPlaca: "+placa if placa else ""
        pnome = "\nNome: "+nome if nome else ""
        pdoc = "\nDoc: "+doc if doc else ""
        Ticket = "Ticket: " + str(num) + \
                 "\nEntrada: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + \
                 pplacas + pnome + pdoc
        open (filename, "w").write (Ticket)
        hinstance = win32api.ShellExecute (
          0,
          "print",
          filename,
          #
          # If this is None, the default printer will
          # be used anyway.
          #
          '/d:"%s"' % win32print.GetDefaultPrinter (),
          ".",
          0
        )
        return True if hinstance > 32 else False


class NovoAloc2(BoxLayout):
    criar_aloc = ObjectProperty()
    cadastrar = ObjectProperty()
    cancel = ObjectProperty()
    user_nome = ObjectProperty()
    user_doc = ObjectProperty()
    user_placas = ObjectProperty()
    data_entrada = ObjectProperty()
    linha_planilha = ObjectProperty()
    nao_cadastrado = ObjectProperty()
    placas_regex = re.compile('(([A-Z]{3}-{1}\d{4}(;|$))|((?=(?:\d*[A-Z]){4})(?=(?:[A-Z]*\d){3})[\w\d]{7}(;|$))){1,}')

    def iter_rows(self, ws):
        row_values = []
        for row in ws.iter_rows():
            yield row#[2].value

    def procurar_cadastro(self, placa):
        placa = placa.upper()
        for row, cliente in enumerate(self.iter_rows(MyApp.user_sheet), 1):
            if placa in str(cliente[2].value):
                return cliente, row
        return False, None

    def valid_placas(self, placa):
        placa = placa.upper()
        result = self.placas_regex.fullmatch(placa)
        return result

    def check_valores(self, placa):
        global MyApp
        if len(placa) > 8:
            self.ids.user_placas.text = placa[0:8]
            return
        self.ids.user_placas.text = self.user_placas.text.upper()
        self.ids.lnome.text = ''
        self.user_nome = ''
        self.ids.ldoc.text = ''
        self.user_doc = ''
        self.nao_cadastrado.text = ''
        self.criar_aloc.disabled = True
        self.label_ver_dados.color = (1,0,0,1)
        self.user_placas.foreground_color = (0, 0, 0, 1)

        if placa and not self.valid_placas(placa) or not placa:
            self.label_ver_dados.text = 'Padrão: AAA-1234.\nPadrão Mercosul: 4 letras e 3 números em qualquer ordem.' 
            self.user_placas.foreground_color = (1, 0, 0, 1)
            return
        else:
            cliente, row = self.procurar_cadastro(placa)
            if cliente:
                data = [cliente[0].value, cliente[1].value, cliente[2].value, cliente[3].value, str(row)]
                self.ids.lnome.text = 'Nome: '
                self.user_nome = cliente[0].value
                self.ids.ldoc.text = 'Doc:    '
                self.user_doc = cliente[1].value
                #print(data)
                self.cadastrar.text = 'Editar'
                self.cadastrar.on_release = partial(MyApp.show_user, data, None)
                self.criar_aloc.disabled = False
            else:
                self.nao_cadastrado.text = 'Não cadastrado.'
                self.cadastrar.text = 'Cadastrar'
                self.cadastrar.on_release = partial(MyApp.adicionar_usuario, self.ids.user_placas.text)
                self.criar_aloc.disabled = False

        self.label_ver_dados.color = (0,1,0,1)
        self.label_ver_dados.text = 'Valores válidos!'

    def adicionar_aloc(self, popup):
        #linha = self.linha_planilha.split(': ')[1]
        try:
            #MyApp.carregar_sheet_aloc()
            row_n=idp.sheet_len(MyApp.aloc_sheet)+1
            if self.user_nome:
                MyApp.aloc_sheet.cell(row=row_n, column=1).value = self.user_nome
            if self.user_doc:
                MyApp.aloc_sheet.cell(row=row_n, column=2).value = self.user_doc
            MyApp.aloc_sheet.cell(row=row_n, column=3).value = self.user_placas.text
            MyApp.aloc_sheet.cell(row=row_n, column=4).value = self.data_entrada.split(':\n')[1]
            #MyApp.aloc_sheet.cell(row=row_n, column=5).value = ''
            MyApp.aloc_sheet.cell(row=row_n, column=6).value = ''
            MyApp.aloc_sheet.cell(row=row_n, column=7).value = 'NÃO'
            MyApp.aloc_sheet.cell(row=row_n, column=8).value = row_n

            success = self.imprimir_ticket(row_n, self.user_nome, self.user_doc, self.user_placas.text)
            MyApp.all_sheets_aloc.save(MyApp.aloc_sheet_path)
            MyApp.load_aloctable_results(self.user_doc)

            pplacas = "\nPlaca: "+self.user_placas.text if self.user_placas.text else ""
            pnome = "\nNome: "+self.user_nome if self.user_nome else ""
            pdoc = "\nDoc: "+self.user_doc if self.user_doc else ""
            Ticket = "Ticket: " + str(row_n) + \
                     "\nEntrada: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + pplacas + pnome + pdoc
            if not success:
                msg = Label(text='Não foi possível imprimir o ticket. Impressora não localizada.\nAnote manualmente os dados:\n\n'+Ticket)
                popup2 = Popup(title='Erro',
                          content=msg,
                          size_hint=(0.5, 0.5))
                popup2.open()
            else:
                msg = Label(text='Comando enviado para impressora! Verifique se há papel.\n\n'+Ticket)
                popup2 = Popup(title='Imprimindo',
                          content=msg,
                          size_hint=(0.5, 0.5))
                popup2.open()
            #MyApp.fechar_sheet_aloc()
        except Exception as e:
            traceback.print_exc(file=sys.stdout)
            msg = Label(text='Impressora não localizada!')
            popup2 = Popup(title='Erro',
                          content=msg,
                          size_hint=(0.3, 0.25))
            popup2.open()
        popup.dismiss()

    def imprimir_ticket(self, num, nome, doc, placa):
        try:
            filename = tempfile.mktemp (".txt")
            pplacas = "\nPlaca: "+placa if placa else ""
            pnome = "\nNome: "+nome if nome else ""
            pdoc = "\nDoc: "+doc if doc else ""
            Ticket = "Ticket: " + str(num) + \
                     "\nEntrada: " + datetime.now().strftime("%d/%m/%Y %H:%M:%S") + \
                     pplacas + pnome + pdoc
            open (filename, "w").write (Ticket)
            hinstance = win32api.ShellExecute (
              0,
              "print",
              filename,
              #
              # If this is None, the default printer will
              # be used anyway.
              #
              '/d:"%s"' % win32print.GetDefaultPrinter (),
              ".",
              0
            )
            return True if hinstance > 32 else False
        except Exception as e:
            return False

class Parking(App):
    def build(self):
        global MyApp
        self.title = 'Estacionamento'
        Builder.load_string(open("UI.kv", encoding="utf-8").read(), rulesonly=True)
        MyApp = UI()
        Window.bind(on_request_close=self.on_request_close)
        return MyApp

    def on_request_close(self, *args):
        #self.backup_planilhas()
        idp.fechar_sheet(MyApp.all_sheets_user)
        idp.fechar_sheet(MyApp.all_sheets_aloc)
        idp.fechar_sheet(MyApp.all_sheets_tb)
        self.stop()
        return True

    def backup_planilhas(self):
        global MyApp
        copy(MyApp.user_sheet_path, 'bk_user.xlsx')
        copy(MyApp.aloc_sheet_path, 'bk_aloc.xlsx')
        

Factory.register('UserDialog', cls=UserDialog)
Factory.register('NovoUser', cls=NovoUser)
Factory.register('AlocDialog', cls=AlocDialog)
Factory.register('NovoAloc', cls=NovoAloc)
Factory.register('Menu', cls=Menu)

if __name__ == '__main__':
    Parking().run()
