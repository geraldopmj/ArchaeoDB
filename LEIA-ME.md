# ArchaeoDB

O ArchaeoDB é um software desenvolvido para facilitar a gestão e análise de dados em zooarqueologia. Ele permite o cadastro, organização, visualização e manipulação de informações sobre vestígios ósseos, auxiliando pesquisadores na estruturação e interpretação de seus dados.

## Documentação

- README em Inglês: [README.md](README.md)  
- Leia-me em Português: [README.pt-BR.md](README.pt-BR.md)

## Principais Funcionalidades

- **Gestão de Dados**: Gerenciamento completo de Sítios, Coleções, Conjuntos (Assemblages), Unidades de Escavação, Níveis, Materiais e Espécimes.  
- **Banco de Dados**: Utiliza SQLite para armazenamento robusto e portátil. Permite criar novos bancos de dados ou abrir bancos existentes com facilidade.  
- **Importação e Exportação**:  
  - Criação e atualização do banco a partir de arquivos Excel.  
  - Exportação completa do banco de dados para Excel.  
  - Geração de relatórios em PDF para Materiais, Unidades e Espécimes.  
- **Estatísticas e Visualização**:  
  - Geração de gráficos por Tipo de Material, Descrição e Quantidade.  
  - Visualização de densidade por Unidade.  
  - Cálculo e plotagem de NISP (Número de Espécimes Identificados) por Táxon.  
  - Exportação de gráficos em formato PNG.  
- **Experiência do Usuário**:  
  - Interface com filtros hierárquicos.  
  - Suporte multilíngue: Inglês, Português e Espanhol.  
  - Temas claro e escuro.  

## Instalação

### 1. Instalar o Python

Baixe e instale o Python 3.10 ou superior em:

https://www.python.org/downloads/

Durante a instalação no Windows, marque a opção:

- "Add Python to PATH"

Após a instalação, verifique no terminal:

```bash
python --version
```

---

### 2. Clonar o Repositório

```bash
git clone https://github.com/seu-usuario/ArchaeoDB.git
cd ArchaeoDB
```

---

### 3. Executar o Setup

No Windows, execute:

```bash
setup.bat
```

O script `setup.bat` irá:

- Criar um ambiente virtual (venv)
- Instalar todas as dependências listadas no `requirements.txt`

Aguarde a conclusão da instalação antes de continuar.

---

## Executar o Programa

Após rodar o `setup.bat`, execute:

```bash
run.bat
```

O script `run.bat` irá:

- Ativar o ambiente virtual
- Iniciar a aplicação principal

Importante:

- É obrigatório executar o `setup.bat` pelo menos uma vez antes de usar o `run.bat`.
- Caso o arquivo `requirements.txt` seja atualizado, execute o `setup.bat` novamente.


## Licença

Este projeto está licenciado sob a **GNU AFFERO GENERAL PUBLIC LICENSE  
Versão 3, 19 de novembro de 2007 (AGPL-3.0)**.

Nos termos da AGPL-3.0:

- É permitido usar, modificar e redistribuir o software.  
- Qualquer versão modificada que seja distribuída ou disponibilizada por meio de rede deve permanecer sob a licença AGPL-3.0.  
- O código-fonte completo correspondente deve ser disponibilizado aos usuários que interajam com o software via rede.  

Consulte o arquivo `LICENSE` para o texto legal completo.

## Isenção de Garantia

Este software é fornecido **"NO ESTADO EM QUE SE ENCONTRA"**, sem qualquer tipo de garantia, expressa ou implícita, incluindo, mas não se limitando a:

- Comercialização  
- Adequação a uma finalidade específica  
- Não violação de direitos  

Os autores e colaboradores não se responsabilizam por quaisquer reivindicações, danos ou outras responsabilidades, seja em contrato, ato ilícito ou de outra natureza, decorrentes do uso ou da impossibilidade de uso do software. O uso é de inteira responsabilidade do usuário.

## Contribuições

Contribuições são bem-vindas sob os termos da licença AGPL-3.0. Ao submeter código, você concorda que sua contribuição será licenciada sob a mesma licença.

## Finalidade

O ArchaeoDB é destinado a fins acadêmicos, científicos e educacionais em zooarqueologia e áreas correlatas. Não substitui normas institucionais de curadoria, políticas de governança de dados ou exigências regulatórias.

## Contato

Para dúvidas, sugestões ou suporte:  
geraldo.pmj@gmail.com

