-- GB/T 9704-2012 党政机关公文格式 - 元数据处理
-- 处理公文元数据，包括发文机关、发文字号、密级等

function Meta(m)
  -- 设置默认值
  if not m.document_type then
    m.document_type = "通知"
  end

  if not m.issuing_authority then
    m.issuing_authority = "某某机关"
  end

  if not m.document_number then
    m.document_number = ""
  end

  if not m.urgency_level then
    m.urgency_level = "普通"
  end

  if not m.confidentiality_level then
    m.confidentiality_level = "公开"
  end

  -- 设置页码格式
  if not m.page_number_format then
    m.page_number_format = "arabic"
  end

  -- 设置标题格式要求
  m.title_font = "小标宋体"
  m.title_size = "22pt"
  m.title_alignment = "center"

  -- 设置正文格式要求
  m.body_font = "仿宋_GB2312"
  m.body_size = "16pt"
  m.line_spacing = "28.95pt"

  -- 设置段落格式
  m.paragraph_indent = "2字符"
  m.paragraph_spacing = "0"

  return m
end