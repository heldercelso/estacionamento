USER: admin
SENHA: admin

----------------------------------------------
Passo a passo simplificado:
----------------------------------------------
(necessário python 3.7.6 com pip):

1. para configurar o projeto completo (virtualenv e gerar .exe) execute gerar_tudo.bat
(faz a etapa 2 seguida da 3)

2. para criar o virtual env execute projeto\gerar_venv.bat
(cria o virtual env - kivy_venv - e ativa, depois instala o requirements.txt)

3. para criar os arquivos executaveis (pasta dist) execute gerar_exe.bat (ainda não é o .exe único)
(ativa o virtual env, roda pyinstaller e deleta arquivos desnecessarios dentro do dist)

4. para debugar com virtual env ativo use projeto\debug.bat

- Rodar programa sem executável:
duplo clique no projeto\Parking.vbs ou abrir debug.bat e digitar python Parking.py

- Rodar programa com executável:
duplo clique em dist\Parking\Parking.exe

5. para gerar instalador único (.exe):
- baixe e instale o Inno Setup Compiler;
- execute o arquivo estacionamento\parkinginno.iss
- a saida deve estar em estacionamento\Output

----------------------------------------------
Passo a passo completo:
----------------------------------------------
Rodar o programa (necessário python 3.7.6 com pip):

- criar virtual env dentro da pasta 'projeto':
1. abrir debug.bat
2. python -m venv kivy_venv

- ativar a virtual env:
1. abra debug.bat novamente
(conteudo: start cmd.exe /k ".\kivy_venv\Scripts\activate.bat")

- instalar requerimentos listados (com virtualenv ativo):
pip install -r requirements.txt

- comando para executar o programa (com virtualenv ativo):
1. no terminal: python Parking.py
2. clique duplo no arquivo: Parking.vbs


----------------------------------------------
Criando arquivos executáveis (pasta dist usando PyInstaller)

<<Seguir tutorial do link:>>
https://kivy.org/doc/stable/guide/packaging-windows.html

Resumo:
---Se já tiver o Parking.spec pular para etapa 4.---

1. dentro da pasta estacionamento execute o comando para gerar o .spec (com virtualenv ativo)
python -m PyInstaller --icon parkingapp.ico projeto\\Parking.py

2. abra o Parking.spec criado e edite:

2.1. adicione no topo:
from kivy_deps import sdl2, glew

2.2. edite o coll:
coll = COLLECT(exe, Tree('projeto\\'),
               a.binaries,
               a.zipfiles,
               a.datas,
               *[Tree(p) for p in (sdl2.dep_bins + glew.dep_bins)],
               strip=False,
               upx=True,
               name='Parking')

2.3. mude o debug para False (caso apresente alguma falha volte para True para verificar)
debug=False

2.4 mude o console para False (caso apresente alguma falha volte para True para verificar)
console=False

4. rode o gerar_exe.bat (esse script ativa a virtualenv, executa o spec e 
por ultimo deleta arquivos desnecessarios)
(comando para rodar apenas o spec: python -m PyInstaller Parking.spec)

5. o exe é gerado em dist/Parking/Parking.exe

----------------------------------------------
Criando o instalador:
1. Baixe e instale o Inno Setup Compiler;
2. Gere o executável seguindo as etapas anteriores (PyInstaller);
3. Abra o arquivo parkinginno.iss, verifique o path para o Parking e compile
4. A saida geralmente fica em Output\parking.exe