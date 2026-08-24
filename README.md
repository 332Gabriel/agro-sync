# Agro-Sync

Sistema de registro de aplicações agrícolas em campo, feito para o pequeno
agricultor. A ideia é ser útil a quem coloca o alimento na mesa — sem exigir
internet, cadastro complicado ou conhecimento técnico.

---

## O problema

Quem aplica defensivo, adubo ou qualquer produto no talhão precisa manter o
registro do que usou, quanto usou e onde. Essa informação é exigida por
fiscalização, por quem compra a produção, e serve de base para o contador
fazer o acompanhamento financeiro.

O problema é que o registro acontece no campo — onde quase nunca há sinal.
Um sistema que só funciona conectado não serve para quem está no meio do
talhão. Por isso o Agro-Sync grava primeiro no aparelho e sincroniza depois,
quando a conexão voltar.

---

## Como funciona

**Registrar uma aplicação:**

1. Anota o produto
2. Informa a quantidade usada por hectare
3. Seleciona o talhão onde foi aplicado

O registro é salvo no aparelho imediatamente, com ou sem internet. Quando a
conexão volta, ele sincroniza sozinho.

**Exportar para o contador:**

4. Escolhe o período (tudo, ou um intervalo de datas)
5. O sistema gera uma planilha formatada, pronta para filtrar e imprimir

---

## Decisões técnicas

### ID gerado no cliente

O identificador de cada registro é um UUID criado no próprio aparelho, antes
de qualquer tentativa de envio.

O motivo: se a rede cair depois que o servidor gravou, mas antes da resposta
chegar, o aparelho não tem como saber se deu certo — então ele reenvia. Como
o registro já vem com identidade própria, o banco reconhece que aquele ID já
existe e ignora a cópia em silêncio.

Se o ID fosse gerado no servidor, cada reenvio criaria um registro novo, e a
mesma aplicação apareceria duas vezes.

### Banco append-only

O banco só aceita inserção. Não existe `UPDATE` nem `DELETE` em nenhum ponto
do sistema.

O motivo é a rastreabilidade. Registro de aplicação pode ser exigido por
fiscalização ou por quem compra a produção — e um registro que pode ser
alterado depois não serve como prova, porque não há como saber se ele é o
original.

Quando algo foi digitado errado, a correção não sobrescreve: entra um registro
novo, com o campo `corrige_id` apontando para o antigo. O errado permanece no
banco, visível. É isso que forma a trilha.

O sistema não prova que a aplicação ocorreu no campo — isso depende de nota
fiscal e outros documentos. O que ele garante é que o registro feito naquele
dia não foi modificado depois.

### O cliente só remove o que o servidor confirmou

O servidor valida cada registro recebido. Se algum campo obrigatório estiver
vazio, ele recusa — o registro não entra no banco de forma nenhuma — e devolve
a lista de recusados com o motivo de cada um.

O aparelho então remove da fila apenas os registros confirmados. Os recusados
permanecem salvos localmente, destacados na tela com o motivo da recusa, até
que o agricultor corrija e sincronize de novo.

Recusa não é descarte. Um registro incompleto ainda representa uma aplicação
que aconteceu no campo — perdê-lo seria pior que mantê-lo pendente.

---

## Como rodar

```bash
pip install -r requirements.txt
python app.py
```

Acesse `http://localhost:5000`

---

## Stack

Flask · PostgreSQL · JavaScript · LocalStorage · openpyxl

---

## Limitações conhecidas

- **Sem autenticação.** Qualquer pessoa com o endereço acessa o sistema.
- **Datas gravadas como texto.** Na planilha exportada, isso impede ordenar e
  filtrar por período de verdade — o Excel trata como string.
- **Campo de produto livre, sem catálogo.** Em teste, "glifosato",
  "glifosfato" e "Glifosato 480" foram registrados como três produtos
  distintos. Qualquer soma por produto sai errada.
- **PostgreSQL em arquivo local.** Não sobrevive a deploy em container sem volume
  persistente configurado.
- **Servidor de desenvolvimento.** O Flask embutido não é adequado para
  produção.

---

## Próximos passos

- Catálogo de produtos, para eliminar as grafias divergentes
- Tratamento correto de data e fuso horário
- Autenticação
- Base de consulta técnica (bulas, carência, janela de plantio) para que o
  sistema deixe de ser só um formulário e passe a avisar o agricultor sobre
  prazos e restrições do produto aplicado

---

## Autor

Wendel Gabriel Riger — [github.com/332Gabriel](https://github.com/332Gabriel)'