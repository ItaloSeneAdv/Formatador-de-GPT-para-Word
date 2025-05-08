Macro Word – FormatadorDeChatGPT

Esta macro foi criada para automatizar a formatação de textos gerados por inteligência artificial, como o ChatGPT. Ela limpa e formata o conteúdo para deixá-lo pronto para uso.

O que esta macro faz

1. Converte texto marcado com dois asteriscos (\*\*) em texto em negrito
   Exemplo: **importante** → importante (em negrito)

2. Remove linhas feitas com três hifens (---)
   Normalmente usadas como divisores manuais ou marcas temporárias

3. Remove símbolos de cerquilha (#) isolados
   Usados por IA ou usuários como marcador de seção ou tópico

4. Remove linhas horizontais automáticas do Word
   Elimina as linhas geradas ao digitar --- e pressionar Enter, que são inseridas como objetos InlineShape

Como funciona

* Opera no documento ativo do Word (ActiveDocument)
* Usa o objeto Range para percorrer o conteúdo
* Utiliza o recurso Find com curingas (wildcards) para identificar padrões como **texto**
* Itera todos os InlineShapes e remove aqueles do tipo wdInlineShapeHorizontalLine

Requisitos

* Microsoft Word com suporte a Macros
* Macros habilitadas no ambiente

Como usar

1. Abra o Word
2. Pressione ALT + F11 para abrir o Editor VBA
3. No menu Inserir > Módulo, cole o código da macro
4. Salve
5. Execute a macro FormatadorDeChatGPT pressionando F5 ou vinculando a um botão personalizado

Exemplo

Antes:

```
**Este trecho precisa de destaque**

---

# Observação importante
```

Depois de executar a macro:

* “Este trecho precisa de destaque” ficará em negrito
* Os separadores (---) e (#) serão removidos
* Linhas horizontais invisíveis no Word serão eliminadas


