-- Converte blocos de codigo ```mermaid em <div class="mermaid"> para o Mermaid JS renderizar
function CodeBlock(el)
  if el.classes and el.classes[1] == "mermaid" then
    return pandoc.RawBlock("html", '<div class="mermaid">\n' .. el.text .. '\n</div>')
  end
end
