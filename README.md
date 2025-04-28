# Sistema de Escalas

Um sistema desktop para gerenciamento de escalas de plantão de funcionários (12x36) e estagiários, com registro de trocas, edição de trocas, atualização mensal e exportação para Excel com destaque de feriados nacionais, estaduais (TO) e municipais (Palmas).

## Funcionalidades

- **Cadastro de Funcionários**: 12x36 (pares/ímpares) e estagiários.
- **Geração de Escala**: automática por mês/ano, via calendário.
- **Troca de Plantão**: registro de trocas, controle de quem substituiu.
- **Edição de Trocas**: alterar data de trocas já realizadas.
- **Atualização de Escala**: limpar e gerar novamente toda a escala de um mês informado.
- **Exportação para Excel**:
  - Planilhas separadas por turno.
  - Coluna extra “Feriado” com nome do feriado.
  - Destaque em cor para dias de feriado.

## Requisitos

- Python 3.x
- Bibliotecas Python:
  - `tkinter` (inclusa no Python)
  - `sqlite3` (inclusa no Python)
  - `openpyxl`
  - `holidays`  

## Instalação

1. Clone este repositório:
   ```bash
   git clone https://github.com/SEU_USUARIO/seu-repo.git
   cd seu-repo
   ```
2. Instale as dependências:
   ```bash
   pip install openpyxl holidays
   ```

## Uso

Execute o script principal:

```bash
python calend.py
```

1. **Cadastro**: preencha nome, tipo, escala, turno e mês/ano.
2. **Troca de Turno**: selecione original, substituto e data.
3. **Editar Troca**: escolha troca existente e nova data.
4. **Atualizar Escala**: informe mês/ano e regenere a escala.
5. **Exportar**: gere a planilha Excel com feriados destacados.

## Estrutura do Banco de Dados

- **funcionarios**: id, nome, tipo, escala_dias, turno
- **escalas**: id, funcionario_id, data, turno, original
- **trocas**: id, data, funcionario_original, funcionario_substituto

## Feriados

- Feriados nacionais e estaduais (TO) obtidos via `holidays.Brazil(subdiv='TO')`.
- Feriados municipais de Palmas definidos no dicionário `palmas_holidays` (formato `"DD/MM": "Nome do Feriado"`).

## Contribuindo

1. Faça um fork deste repositório
2. Crie uma branch com a feature ou correção (`git checkout -b feature/nome`)
3. Faça commit das suas mudanças (`git commit -m "Adiciona feature X"`)
4. Dê push na branch (`git push origin feature/nome`)
5. Abra um Pull Request

## Licença

Distribuído sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

