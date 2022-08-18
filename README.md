
# Introdução

Software para gerenciamento de estacionamentos.
Usuário e senha: admin

Imagens: https://helder-portfolio.herokuapp.com/parking-software-3/


# Tecnologias

 - Python 3.7.6
 - Kivy 1.11.1 (UI)
 - PyInstaller 3.6 (Gerar executáveis)
 - Inno Setup Compiler (Gerar instalador)


# Passo a passo simplificado:

1. Para configurar o projeto completo (virtualenv e gerar arquivos Pyinstaller) execute o arquivo `gen_all.bat`.

    ATENÇÃO: Isso executará as etapas 1.1 e 1.2 abaixo automaticamente.

    1.1. Se desejar criar somente a virtualenv execute o arquivo `projeto/gen_venv.bat`.

        Etapas: Cria a pasta `kivy_venv` e instala o `requirements.txt`.

    1.2. Se dejesar criar somente os arquivos do Pyinstaller (pasta dist) execute `gen_exe.bat` (ainda não é o .exe único).

        Etapas: Ativa a virtualenv, executa o Pyinstaller e deleta arquivos desnecessarios dentro do dist

OBS: Para abrir terminal com virtualenv já ativo use `projeto/debug.bat`

2. Para executar o software:

    - Rodar programa sem executável:
    Duplo clique no `projeto/Parking.vbs` ou abrir `debug.bat` e digitar `python Parking.py`.

    - Rodar programa com executável:
    Duplo clique em `dist/Parking/Parking.exe`

3. Para gerar instalador único (.exe):

    * Baixe e instale o `Inno Setup Compiler`;
    * Execute o arquivo `estacionamento/parkinginno.iss`;
    * A saída deve estar em `estacionamento/Output`.


# Passo a passo detalhado:

* Criar virtualenv dentro da pasta 'projeto':

    1. Abrir `debug.bat` (conteudo: `start cmd.exe /k ".\kivy_venv\Scripts\activate.bat"`);

    2. Executar `python -m venv kivy_venv`.

* Ativando a virtualenv e instalando libs:

    1. Abrir `debug.bat`;

    2. Instalar libs com o comando `pip install -r requirements.txt`.

* Modos de executar o programa (com virtualenv ativo):

    1. Abra o terminal (`debug.bat`) e execute o comando `python Parking.py`;

    2. Clique duplo no arquivo `Parking.vbs`.


# Criando arquivos executáveis (PyInstaller)

* Tutorial completo no link:

`https://kivy.org/doc/stable/guide/packaging-windows.html`

* Resumo:

Se já tiver o `Parking.spec` pular para `etapa 4`.

1. Dentro da pasta estacionamento execute o comando para gerar o `.spec` (com virtualenv ativo):

    ``` shell
    $ python -m PyInstaller --icon parkingapp.ico projeto\\Parking.py
    ```

2. Abra o `Parking.spec` criado e edite:

    2.1. Adicione no topo:

        `from kivy_deps import sdl2, glew`

    2.2. Edite a variável `coll`:

        ```
        coll = COLLECT(exe, Tree('projeto\\'),
                       a.binaries,
                       a.zipfiles,
                       a.datas,
                       *[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins)],
                       strip=False,
                       upx=True,
                       name='Parking')
        ```

    2.3. Na variável `exe=...`, mude o `debug` para `False` (caso apresente alguma falha volte para True para verificar)

    2.4. Na variável `exe=...`, mude o `console` para `False` (caso apresente alguma falha volte para True para verificar)


3. Execute o arquivo `gen_exe.bat` (esse script ativa a virtualenv, executa o spec e por último deleta arquivos desnecessários)

    -> Se desejar executar apenas o spec:

        ``` shell
        $ python -m PyInstaller Parking.spec)
        ```

5. O `.exe` é gerado em `dist/Parking/Parking.exe`


# Criando o instalador

1. Baixe e instale o `Inno Setup Compiler`;
2. Gere o executável seguindo as etapas anteriores (PyInstaller);
3. Abra o arquivo parkinginno.iss, verifique o path para o Parking e compile;
4. A saída geralmente fica em `Output/parking.exe`.