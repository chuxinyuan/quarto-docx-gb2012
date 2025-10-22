-- GB/T 9704-2012 党政机关公文格式 - 行号处理
-- 为 Word 文档添加行号（可选功能）

function Docx(doc)
  local meta = doc.meta

  -- 检查是否需要添加行号
  if meta.add_line_numbers and meta.add_line_numbers == true then
    -- 添加行号设置到文档属性
    return addLineNumbers(doc)
  end

  return doc
end

-- 添加行号到文档
function addLineNumbers(doc)
  -- 为文档添加行号设置的XML
  local line_number_xml =
    '<w:sectPr><w:lnNumType w:countBy="1" w:start="1" w:restart="continuous"/></w:sectPr>'

  -- 将行号设置添加到文档末尾
  local line_number_block = pandoc.RawBlock("openxml", line_number_xml)
  table.insert(doc.blocks, line_number_block)

  return doc
end

-- 为段落添加行号标记
function addLineNumberToParagraph(para, line_number)
  local line_number_span = pandoc.Span({pandoc.Str(tostring(line_number))})
  line_number_span.attributes = {class = "line-number", style = "float: left; width: 2em; text-align: right;"}

  -- 在段落前添加行号
  if para.content then
    table.insert(para.content, 1, line_number_span)
    table.insert(para.content, 2, pandoc.Space())
  end

  return para
end