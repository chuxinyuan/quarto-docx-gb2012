-- GB/T 9704-2012 党政机关公文格式 - Word样式处理
-- 为生成的Word文档应用符合国标的样式

function Docx(doc)
  -- 添加样式定义到Word文档
  local styles = {
    -- 公文标题样式：小标宋体，22pt，红色
    'GBDocumentTitle',
    -- 发文机关标识样式：宋体，14pt，红色
    'GBIssuingAuthority',
    -- 发文字号样式：仿宋，16pt
    'GBDocumentNumber',
    -- 正文一级标题样式：黑体，16pt
    'GBLevel2Title',
    -- 正文二级标题样式：楷体，16pt
    'GBLevel3Title',
    -- 正文段落样式：仿宋，16pt，首行缩进2字符
    'GBParagraph',
    -- 文档标识样式（密级、紧急程度）：楷体，16pt
    'GBIdentifier',
    -- 签发人样式：楷体，16pt
    'GBSignatory',
    -- 主送机关样式：仿宋，16pt
    'GBRecipients',
    -- 列表项样式：仿宋，16pt
    'GBListItem',
    -- 页眉样式：宋体，14pt，居中
    'GBHeader',
    -- 页脚样式：宋体，10.5pt，居中
    'GBFooter'
  }

  return doc
end

-- 处理文档的样式应用
function applyStyles(block)
  if block.attributes and block.attributes.style then
    local style_name = block.attributes.style

    -- 根据样式名称应用相应的格式
    if style_name == "GBDocumentTitle" then
      return applyTitleStyle(block)
    elseif style_name == "GBParagraph" then
      return applyParagraphStyle(block)
    elseif style_name == "GBIssuingAuthority" then
      return applyIssuingAuthorityStyle(block)
    elseif style_name == "GBDocumentNumber" then
      return applyDocumentNumberStyle(block)
    elseif style_name == "GBIdentifier" then
      return applyIdentifierStyle(block)
    elseif string.match(style_name, "GBLevel%d+Title") then
      return applyTitleLevelStyle(block, style_name)
    end
  end

  return block
end

-- 应用公文标题样式
function applyTitleStyle(block)
  -- 小标宋体，22pt，红色，居中
  block = pandoc.RawBlock("openxml",
    '<w:p><w:pPr><w:pStyle w:val="GBDocumentTitle"/></w:pPr>' ..
    '<w:r><w:rPr><w:rFonts w:ascii="小标宋体" w:eastAsia="小标宋体"/>' ..
    '<w:sz w:val="44"/><w:szCs w:val="44"/><w:color w:val="FF0000"/>' ..
    '</w:rPr><w:t>' .. extractText(block) .. '</w:t></w:r></w:p>')

  return block
end

-- 应用正文段落样式
function applyParagraphStyle(block)
  -- 仿宋，16pt，首行缩进2字符，行距28.95pt
  block = pandoc.RawBlock("openxml",
    '<w:p><w:pPr><w:pStyle w:val="GBParagraph"/><w:ind w:firstLine="720"/>' ..
    '<w:spacing w:line="579" w:lineRule="exact"/></w:pPr>' ..
    '<w:r><w:rPr><w:rFonts w:ascii="仿宋_GB2312" w:eastAsia="仿宋_GB2312"/>' ..
    '<w:sz w:val="32"/><w:szCs w:val="32"/></w:rPr>' ..
    '<w:t>' .. extractText(block) .. '</w:t></w:r></w:p>')

  return block
end

-- 应用发文机关标识样式
function applyIssuingAuthorityStyle(block)
  -- 宋体，14pt，红色，居中
  block = pandoc.RawBlock("openxml",
    '<w:p><w:pPr><w:pStyle w:val="GBIssuingAuthority"/><w:jc w:val="center"/></w:pPr>' ..
    '<w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体"/>' ..
    '<w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="FF0000"/>' ..
    '</w:rPr><w:t>' .. extractText(block) .. '</w:t></w:r></w:p>')

  return block
end

-- 提取文本内容的辅助函数
function extractText(block)
  if block.content then
    local text = ""
    for i, element in ipairs(block.content) do
      if element.text then
        text = text .. element.text
      elseif element.c then
        text = text .. extractTextFromInline(element)
      end
    end
    return text
  end
  return ""
end

-- 从内联元素中提取文本
function extractTextFromInline(element)
  if element.text then
    return element.text
  elseif element.c then
    local text = ""
    for i, child in ipairs(element.c) do
      text = text .. extractTextFromInline(child)
    end
    return text
  end
  return ""
end