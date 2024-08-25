from pptx import Presentation
import PySimpleGUI as sg

class Pptxautomation:

    def __init__(self):
        prs = Presentation('PADRAO-CULTO-ONLINE.pptx')
        
        tema = self.escolhe_tema()
        
        titulo, pregador = self.define_titulo_e_pregador(prs)

        slide = prs.slides.add_slide(prs.slide_masters[1].slide_layouts[0])
        texto  = slide.shapes.title
        # texto.text = f'{titulo} {pregador}'
        texto.text = titulo.strip()

        self.change_pregador_name_to_bold(prs, pregador.strip())

        # Adicionando texto chave
        self.define_versiculos(prs, 1)
        
        pontos = self.define_pontos(prs, tema)

        prs.save(f'{titulo.strip()} - {pregador.strip()}.pptx')

    def escolhe_tema(self):
        sg.theme('DarkAmber')
        tema = 1

        layout = [
            [sg.Text(f'Qual Tema deseja usar?')],
            [sg.Text(f'Obs: Tema1: Pontos em laranja, Tema2: Pontos em branco')],
            [sg.Button('Tema1'), sg.Button('Tema2')],
        ]

        window = sg.Window('Escolha de tema', layout)
        event, values = window.read()

        window.close()

        match event:
            case 'Tema1':
                tema = 1
            case 'Tema2':
                tema = 2

        return tema
  
    def define_titulo_e_pregador(self, presentation):
        sg.theme('DarkAmber')

        layout = [
            # [sg.Output(size=(30,0), key='respostas')],
            [sg.Text('Digite o título da palavra')],
            [sg.Input(default_text='', key='titulo')],
            [sg.Text('Digite o nome do pregador da palavra')],
            [sg.Input(default_text='', key='pregador')],
            [sg.Button('Confirmar')],
        ]

        window = sg.Window('Automação', layout)

        event, values = window.read()

        titulo = values['titulo']
        pregador = values['pregador']
        window.close()

        return titulo.title(), pregador.title()
    
    def define_pontos(self, presentation, tema):
        sg.theme('DarkAmber')

        i = 1
        flag = 1
        pontos = []

        while flag == 1: # Enquanto usuário quiser adicionar pontos o while continuará

            # Criando tela que pega o nome do ponto
            layout = [
                [sg.Text(f'Digite o texto do ponto {i}')],
                [sg.Input(default_text='', key='ponto')],
                [sg.Button('Confirmar')],
            ]

            window = sg.Window('Automação', layout)

            event, values = window.read()

            window.close()

            ponto = values['ponto']
            pontos.append(ponto)

            self.cria_slide_de_ponto(presentation, ponto, i, tema)

            # Adicionando versiculos ao pprt
            self.define_versiculos(presentation, 2)

            layout = [
                [sg.Text(f'Deseja adicionar outro ponto?')],
                [sg.Button('Sim'), sg.Button('Não')],
            ]

            window = sg.Window('Automação', layout)

            event, values = window.read()
            window.close()

            match event:
                case 'Sim':
                    flag = 1
                    i = i + 1
                case 'Não':
                    flag = 0
        
        return pontos

    def cria_slide_de_ponto(self, presentation, ponto, nmrPonto, tema):
        slide = presentation.slides.add_slide(presentation.slide_masters[tema].slide_layouts[nmrPonto + 1])
        # subtitulo = slide.placeholders[1]
        subtitulo  = slide.shapes.title
        subtitulo.text = ponto.strip()

    def define_versiculos(self, presentation, mode):
        temVersiculos = False
        flag = True
        
        # Criando tela que pergunta se deseja adicionar versículos ao ponto
        layout = [
            [sg.Text(f'Deseja inserir um texto chave?' if mode == 1 else f'Deseja adicionar versículos a esse ponto?')],
            [sg.Button('Sim'), sg.Button('Não')],
        ]

        window = sg.Window('Automação', layout)

        event, values = window.read()

        window.close()

        match event:
            case 'Sim':
                temVersiculos = True
            case 'Não':
                temVersiculos = False
                return
        

        if temVersiculos == True:
            while flag == True:
                # Criando tela que pega o versículo
                layout = [
                    [sg.Text(f'Digite o versículo que deseja adicionar')],
                    [sg.Input(default_text='', key='versiculosTitle')],
                    [sg.Text(f'Digite o texto dos versículos')],
                    # [sg.Input(default_text='', key='versiculosText')],
                    [sg.Multiline(default_text='', key='versiculosText', size=(None, 5))],
                    [sg.Button('Confirmar'), sg.Button('Cancelar')],
                ]

                window = sg.Window('Automação', layout)

                event, values = window.read()
                window.close()

                versiculosTitle, versiculosText = values['versiculosTitle'], values['versiculosText']

                self.cria_slides_de_versiculo(presentation, versiculosTitle, versiculosText)

                # Criando tela que pergunta se deseja adicionar versículos ao ponto
                layout = [
                    [sg.Text('Deseja inserir outro texto?' if mode == 1 else 'Deseja adicionar outro versículo a esse ponto?')],
                    [sg.Button('Sim'), sg.Button('Não')],
                ]

                window = sg.Window('Automação', layout)

                event, values = window.read()
                window.close()

                match event:
                    case 'Sim':
                        flag = True
                    case 'Não':
                        flag = False
                        return               

    def cria_slides_de_versiculo(self, presentation, versiculoTitle, versiculoText):
        versiculos = versiculoText.splitlines()
        
        for versiculo in versiculos:
            slide = presentation.slides.add_slide(presentation.slide_masters[1].slide_layouts[1])

            titleVer = slide.shapes.title
            titleVer.text = versiculoTitle

            textVer = slide.placeholders[1]
            textVer.text = versiculo
    
    def change_pregador_name_to_bold(self, presentation, pregador):
        slide = presentation.slides[(0)]
        title = slide.shapes.title
        tf = title.text_frame

        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f' {pregador}'
        font = run.font
        font.bold = True

    def update_text_of_textbox(self, presentation, slide, text_box_id, new_text): 
        slide = presentation.slides[(slide - 1)]
        count = 0
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text:
                count += 1
                if count == text_box_id:
                    text_frame = shape.text_frame
                    first_paragraph = text_frame.paragraphs[0]
                    first_run = first_paragraph.runs[0] if first_paragraph.runs else first_paragraph.add_run()
                    # Preserve formatting of the first run
                    font = first_run.font
                    font_name = font.name
                    font_size = font.size
                    font_bold = font.bold
                    font_italic = font.italic
                    font_underline = font.underline
                    # font_color = font.color.rgb
                    # Clear existing text and apply new text with preserved formatting
                    text_frame.clear()  # Clears all text and formatting
                    new_run = text_frame.paragraphs[0].add_run()  # New run in first paragraph
                    new_run.text = new_text
                    # Reapply formatting
                    new_run.font.name = font_name
                    new_run.font.size = font_size
                    new_run.font.bold = font_bold
                    new_run.font.italic = font_italic
                    new_run.font.underline = font_underline
                    # new_run.font.color.rgb = font_color
                    return

Pptxautomation()