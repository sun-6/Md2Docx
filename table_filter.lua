-- Md2Docx — pandoc Lua 过滤器
-- Copyright (C) 2025-2026 sun_yy
-- SPDX-License-Identifier: GPL-3.0-or-later

-- ============================================================
-- 0. 诊断输出与跨语言契约
--    诊断经 stderr 前缀协议回传核心层（md2docx_core.py 按前缀解析）：
--    WARN 升级进转换 warnings，INFO 仅进日志。前缀字符串与核心层
--    两侧硬编码、必须同步修改。
-- ============================================================

local function ref_diag_info(msg)
  io.stderr:write("MD2DOCX-REF-INFO: " .. msg .. "\n")
end

local function ref_diag_warn(msg)
  io.stderr:write("MD2DOCX-REF-WARN: " .. msg .. "\n")
end

-- 以下两条与 core 侧互为镜像（Lua/Python 各一份，改动须双侧同步）：
-- * 引用假值清单 ↔ core convert_md_to_format 的 ("false","no","off","0")
-- * 引用 token 语法   ↔ core 的 _CITATION_TOKEN_RE（编号/逗号/区间分隔符）
-- 引用章节标题识别    ↔ core 的 _REFERENCE_HEADING_RE（见 normalize_heading_text）

-- 前置声明：定义在后文、此前需被引用的实现
local assign_column_widths  -- 3. 表格列宽归一化（HTML 重组表也要归一）
local table_rows            -- 表格行收集（引用递归与列宽计算共用）

-- ============================================================
-- 1. RawBlock HTML 表格 -> 原生 Table
--    markdown reader 会把 HTML 表格拆成 RawBlock/Plain 碎片，
--    docx writer 会丢弃 RawBlock；这里缓冲重组后用 HTML reader
--    重新解析，colspan/class/style 等属性均可保留。
-- ============================================================

local function is_table_start(text)
  return text:match("^%s*<table[%s>]") or text:match("^%s*<table$")
end

local function is_table_end(text)
  return text:match("</table>%s*$")
end

local function wrap_html_table(html)
  -- tex_math_single_backslash: 缓冲回写的 Plain 块中公式被序列化为
  -- <span class="math inline">\(x\)</span>，纯 html reader 不识别会丢内容，
  -- 加该扩展后可还原为 Math 元素
  local ok, doc = pcall(
    pandoc.read,
    "<!doctype html><html><body>" .. html .. "</body></html>",
    "html+tex_math_single_backslash"
  )
  if ok and doc and doc.blocks and #doc.blocks > 0 then
    return doc.blocks
  end
  -- 解析失败降级为 RawBlock，而 docx writer 会丢弃 RawBlock——
  -- 整张表从输出中消失，必须留下诊断，不能无声
  ref_diag_warn("HTML 表格重解析失败或结果为空，该表可能从输出中丢失: " .. html:sub(1, 80):gsub("%s+", " "))
  return pandoc.RawBlock("html", html)
end

local function math_to_raw_literal(block)
  -- pandoc 3.x html writer 会把 Math 渲染成样式化 HTML（<em>E</em> 等），
  -- 缓冲回写后无法还原为公式；先降级为 \(...\) / \[...\] 字面量，
  -- 配合 wrap_html_table 里的 tex_math_single_backslash 读回 Math 元素
  return block:walk({
    Math = function(el)
      if el.mathtype == pandoc.DisplayMath then
        return pandoc.RawInline("html", "\\[" .. el.text .. "\\]")
      end
      return pandoc.RawInline("html", "\\(" .. el.text .. "\\)")
    end,
  })
end

-- 表格行收集：pandoc 3.8 Lua 的 head 带 .rows；bodies[i].head / body
-- 则直接是行列表（无 .rows 字段），两种形态都要收集。
-- 引用递归（table_each_row）与列宽计算共用，行形态处理只此一份。
function table_rows(tbl)
  local rows = {}
  local function collect(part)
    if not part then
      return
    end
    if part.rows then
      for _, row in ipairs(part.rows) do
        rows[#rows + 1] = row
      end
    else
      for _, row in ipairs(part) do
        if pandoc.utils.type(row) == 'Row' then
          rows[#rows + 1] = row
        end
      end
    end
  end
  collect(tbl.head)
  for _, body in ipairs(tbl.bodies or {}) do
    collect(body.head)
    collect(body.body)
  end
  return rows
end

local function table_each_row(tbl, fn)
  for _, row in ipairs(table_rows(tbl)) do
    fn(row)
  end
end

-- ============================================================
-- 2. 引用角标 -> 参考文献交叉引用
--    正文中的 [1] / [2,3] / [1-3] 括号编号角标，当编号能在本文档
--    参考文献章节的编号条目中全部解析时，改写为指向条目书签
--    (ref-N) 的上标内部链接——docx 输出即 Word 交叉引用
--    （w:hyperlink w:anchor + w:bookmarkStart）。
--    判据是语义的（编号可解析才转换），识别在 AST 层完成：
--    行内代码 / 公式 / 链接文字 / 围栏内容天然不受影响。
--    开关：pandoc 元数据 md2docx-citations（核心层经 -M 注入，
--    缺省开启）；typst 等书签语义未验证的 writer 不启用。
-- ============================================================

local REF_HEADING_TEXTS = {
  ["参考文献"] = true,
  ["references"] = true,
  ["bibliography"] = true,
}
local REF_ANCHOR_PREFIX = "ref-"

-- 参考文献标题的前导编号：阿拉伯数字或中文数字，可跟至多一个分隔符
local REF_HEADING_SEPARATORS = { ".", "、", ")", "．" }
local REF_HEADING_NUMERALS = {
  ["一"] = true, ["二"] = true, ["三"] = true, ["四"] = true, ["五"] = true,
  ["六"] = true, ["七"] = true, ["八"] = true, ["九"] = true, ["十"] = true, ["百"] = true,
}

-- 标题文本归一：去首尾空白、剥章节前导编号、英文小写。
-- 编号口径与 core 的 _REFERENCE_HEADING_RE 保持同一（阿拉伯数字 |
-- 中文数字）+ 至多一个分隔符（. 、 ) ．）。中文数字须按真实字符
-- 匹配——Lua 字符类按字节工作，会误伤共用首字节的汉字（如"人"）。
local function normalize_heading_text(text)
  local t = text:gsub("^%s+", ""):gsub("%s+$", ""):lower()

  local function strip_numbering(s)
    local digits = s:match("^(%d+)")
    local num_len = digits and #digits or nil
    if not num_len then
      -- 中文数字前缀：逐字符匹配（utf8 库按码点解码）
      local ok = pcall(function()
        local offset = 0
        for _, codepoint in utf8.codes(s) do
          if not REF_HEADING_NUMERALS[utf8.char(codepoint)] then
            break
          end
          offset = offset + #utf8.char(codepoint)
        end
        if offset > 0 then
          num_len = offset
        end
      end)
      if not ok then
        return s
      end
      if not num_len then
        return s
      end
    end
    local rest = s:sub(num_len + 1)
    for _, sep in ipairs(REF_HEADING_SEPARATORS) do
      if rest:sub(1, #sep) == sep then
        rest = rest:sub(#sep + 1)
        break
      end
    end
    return rest:gsub("^%s+", "")
  end

  return strip_numbering(t)
end

local function citations_enabled(meta)
  -- FORMAT 与 --to 名称对应：html5 输出时是 "html5"，故按前缀匹配。
  -- 假值清单与 core convert_md_to_format 的 ("false","no","off","0") 互为镜像
  local fmt = tostring(FORMAT or "")
  if fmt ~= "docx" and fmt:sub(1, 4) ~= "html" then
    return false
  end
  local flag = meta["md2docx-citations"]
  if flag == nil then
    return true
  end
  if type(flag) == "boolean" then
    return flag
  end
  local s = pandoc.utils.stringify(flag):lower()
  return s ~= "false" and s ~= "off" and s ~= "no" and s ~= "0"
end

local function is_reference_header(header)
  if header.t ~= "Header" then
    return false
  end
  return REF_HEADING_TEXTS[normalize_heading_text(pandoc.utils.stringify(header))] == true
end

-- 引用 token 解析：括号内为逗号分隔的编号 / 编号区间（- – — ~），返回编号列表
local function parse_citation_numbers(inner)
  -- 先把全角/宽分隔符归一到 ASCII，再做字节安全的纯数字校验
  local normalized = inner:gsub("，", ","):gsub("–", "-"):gsub("—", "-"):gsub("~", "-"):gsub("%s+", "")
  local numbers = {}
  for part in normalized:gmatch("[^,]+") do
    local a, b = part:match("^(%d+)%-(%d+)$")
    if a then
      a, b = tonumber(a), tonumber(b)
      -- 区间上限只是防误判护栏（避免把普通数字串当超大区间展开），
      -- 真实文献列表可达数百条（如 [100-150]），不能卡在两位数
      if a > b or b - a > 999 then
        return nil
      end
      for k = a, b do
        numbers[#numbers + 1] = k
      end
    else
      local n = tonumber(part:match("^(%d+)$"))
      if not n then
        return nil
      end
      numbers[#numbers + 1] = n
    end
  end
  if #numbers == 0 then
    return nil
  end
  return numbers
end

-- 在 Str 文本内扫描括号 token，可解析的替换为链接（wrap=true 时再包
-- Superscript；已处于原生 Superscript 语境时传 false，避免嵌套双上标）
local function rewrite_str_tokens(text, entries, stats, wrap)
  if not text:find("[", 1, true) then
    return nil
  end
  local out = pandoc.List()
  local changed = false
  local pos, len = 1, #text
  while pos <= len do
    local s, e = text:find("%b[]", pos)
    if not s then
      break
    end
    local token = text:sub(s, e)
    local numbers = parse_citation_numbers(token:sub(2, -2))
    local resolvable = numbers ~= nil
    if resolvable then
      for _, n in ipairs(numbers) do
        if not entries[n] then
          resolvable = false
        end
      end
    end
    if s > pos then
      out:insert(pandoc.Str(text:sub(pos, s - 1)))
    end
    if resolvable then
      local link = pandoc.Link(
        pandoc.List({ pandoc.Str(token) }),
        "#" .. REF_ANCHOR_PREFIX .. numbers[1],
        "",
        pandoc.Attr("", { "md2docx-ref" })
      )
      out:insert(wrap and pandoc.Superscript(pandoc.List({ link })) or link)
      changed = true
      stats.linked = stats.linked + 1
    else
      -- 形似引用但编号缺失：仅在存在参考文献条目时才值得告警
      if numbers and stats.has_entries then
        for _, n in ipairs(numbers) do
          if not entries[n] then
            stats.dangling[n] = true
          end
        end
      end
      out:insert(pandoc.Str(token))
    end
    pos = e + 1
  end
  if not changed then
    return nil
  end
  if pos <= len then
    out:insert(pandoc.Str(text:sub(pos)))
  end
  return out
end

local rewrite_inlines -- 前向声明：与容器/脚注递归互引
local rewrite_blocks

-- 原地递归单个 inline；Str 的拆分由列表层完成。
-- 注意执行顺序：本过滤器第 2 节的 Inlines handler（自底向上）先于本节
-- 的 Pandoc handler 执行，<sup>[1]</sup> 在进入本节前已是原生
-- Superscript 元素——因此"已在上标语境"的判断针对原生元素，
-- 区间内传 wrap=false 只发裸链接，避免 Superscript(Superscript(Link))。
local function rewrite_inline_element(el, entries, stats, wrap)
  if el.t == "Note" then
    rewrite_blocks(el.content, entries, stats)
  elseif el.t == "Superscript" then
    local new_content = rewrite_inlines(el.content, entries, stats, false)
    if new_content then
      el.content = new_content
    end
  elseif el.t == "Emph" or el.t == "Strong" or el.t == "Underline" or el.t == "Strikeout"
      or el.t == "Subscript" or el.t == "SmallCaps"
      or el.t == "Span" or el.t == "Quoted" then
    local new_content = rewrite_inlines(el.content, entries, stats, wrap)
    if new_content then
      el.content = new_content
    end
  end
  -- Link / Image / Code / Math / RawInline / Cite：内容不参与识别
end

rewrite_inlines = function(inlines, entries, stats, wrap)
  local out = pandoc.List()
  local changed = false
  local function emit(el, inner_wrap)
    if el.t == "Str" then
      local repl = rewrite_str_tokens(el.text, entries, stats, inner_wrap)
      if repl then
        changed = true
        out:extend(repl)
        return
      end
    else
      rewrite_inline_element(el, entries, stats, inner_wrap)
    end
    out:insert(el)
  end
  for i = 1, #inlines do
    emit(inlines[i], wrap)
  end
  if changed then
    return out
  end
  return nil
end

-- 原地递归块：所有改写都发生在块/单元的 content 列表上，块本身身份不变
rewrite_blocks = function(blocks, entries, stats)
  for _, block in ipairs(blocks) do
    local t = block.t
    if t == "Para" or t == "Plain" or t == "Header" then
      local new_content = rewrite_inlines(block.content, entries, stats, true)
      if new_content then
        block.content = new_content
      end
    elseif t == "BlockQuote" or t == "Div" or t == "Figure" then
      rewrite_blocks(block.content, entries, stats)
    elseif t == "BulletList" or t == "OrderedList" then
      for _, item in ipairs(block.content) do
        rewrite_blocks(item, entries, stats)
      end
    elseif t == "Table" then
      table_each_row(block, function(row)
        for _, cell in ipairs(row.cells) do
          rewrite_blocks(cell.contents, entries, stats)
        end
      end)
    end
  end
end

-- 块文本是否以 [n] 开头（参考文献条目的编号前缀）
local function entry_number_of(block)
  local n = pandoc.utils.stringify(block):match("^%s*%[(%d+)%]")
  return n and tonumber(n) or nil
end

-- 条目块首插入空 Span 书签锚点（docx writer 输出 w:bookmarkStart）
local function anchor_block(block, number)
  if block.t ~= "Para" and block.t ~= "Plain" then
    return false
  end
  block.content:insert(1, pandoc.Span(pandoc.List({}), pandoc.Attr(REF_ANCHOR_PREFIX .. number)))
  return true
end

-- 逐行相连的条目段（单 Para 含 SoftBreak）按行拆分成多个条目 Para
local function split_paragraph_entries(block)
  local groups = {}
  local cur = pandoc.List()
  for _, el in ipairs(block.content) do
    if el.t == "SoftBreak" or el.t == "LineBreak" then
      groups[#groups + 1] = cur
      cur = pandoc.List()
    else
      cur:insert(el)
    end
  end
  groups[#groups + 1] = cur
  if #groups < 2 then
    return nil
  end
  -- 保守拆分：每一行都以 [n] 开头才拆（混排时整体保留，仅锚定首行编号）
  for _, g in ipairs(groups) do
    if not pandoc.utils.stringify(pandoc.Plain(g)):match("^%s*%[(%d+)%]") then
      return nil
    end
  end
  return groups
end

-- 两遍式主流程：A 定位章节并锚定条目 -> B 章节外改写角标
local function apply_citation_pass(blocks, meta)
  if not citations_enabled(meta) then
    return blocks
  end

  local entries = {}
  local stats = { linked = 0, dangling = {}, anchored = 0, has_entries = false, duplicates = {} }
  local out = pandoc.List()
  local section_start, section_end
  local in_section = false
  local section_level = 0
  local section_seen = false
  local extra_sections = 0

  local function register(number)
    if entries[number] then
      stats.duplicates[#stats.duplicates + 1] = number
      return false
    end
    entries[number] = true
    stats.anchored = stats.anchored + 1
    stats.has_entries = true
    return true
  end

  local function anchor_if_numbered(block)
    local number = entry_number_of(block)
    if number and register(number) then
      anchor_block(block, number)
    end
  end

  -- Pass A：参考文献章节内的块 -> 编号条目锚定
  local function collect_entry(block)
    local t = block.t
    if t == "Para" or t == "Plain" then
      local groups = split_paragraph_entries(block)
      if groups then
        for _, g in ipairs(groups) do
          local entry_para = pandoc.Para(g)
          anchor_if_numbered(entry_para)
          out:insert(entry_para)
        end
      else
        anchor_if_numbered(block)
        out:insert(block)
      end
    elseif t == "OrderedList" then
      local start = (block.listAttributes and block.listAttributes.start) or 1
      local index = 0
      for _, item in ipairs(block.content) do
        index = index + 1
        local first = item[1]
        if first and (first.t == "Para" or first.t == "Plain") and register(start + index - 1) then
          anchor_block(first, start + index - 1)
        end
      end
      out:insert(block)
    elseif t == "BulletList" then
      for _, item in ipairs(block.content) do
        local first = item[1]
        if first and (first.t == "Para" or first.t == "Plain") then
          anchor_if_numbered(first)
        end
      end
      out:insert(block)
    else
      out:insert(block)
    end
  end

  for _, block in ipairs(blocks) do
    if block.t == "Header" then
      if in_section and block.level <= section_level then
        in_section = false
        section_end = #out
      end
      if is_reference_header(block) then
        if section_seen then
          extra_sections = extra_sections + 1
          out:insert(block)
        else
          section_seen = true
          in_section = true
          section_level = block.level
          out:insert(block)
          section_start = #out + 1
        end
      else
        out:insert(block)
      end
    elseif in_section then
      collect_entry(block)
    else
      out:insert(block)
    end
  end
  if in_section then
    section_end = #out
  end

  -- Pass B：章节范围外的所有块做角标改写
  for index, block in ipairs(out) do
    local in_section_span = section_start ~= nil
      and index >= section_start
      and section_end ~= nil
      and index <= section_end
    if not in_section_span then
      rewrite_blocks({ block }, entries, stats)
    end
  end

  if section_seen then
    if #stats.duplicates > 0 then
      local parts = {}
      for _, n in ipairs(stats.duplicates) do
        parts[#parts + 1] = "[" .. n .. "]"
      end
      ref_diag_warn("参考文献条目编号重复，仅首个条目生成锚点: " .. table.concat(parts, " "))
    end
    if stats.anchored == 0 then
      ref_diag_warn("参考文献章节未检测到 [n] 编号条目，引用角标将保持纯文本")
    elseif stats.linked > 0 then
      ref_diag_info(
        string.format("已将 %d 处引用角标转换为参考文献交叉引用（%d 条参考文献）", stats.linked, stats.anchored)
      )
    else
      ref_diag_info(string.format("参考文献已锚定 %d 条，但未发现正文引用角标", stats.anchored))
    end
  end
  if extra_sections > 0 then
    ref_diag_warn(string.format("检测到 %d 个参考文献章节，仅第一个章节的条目生成锚点", extra_sections + 1))
  end
  local dangling_nums = {}
  for n in pairs(stats.dangling) do
    dangling_nums[#dangling_nums + 1] = n
  end
  if #dangling_nums > 0 then
    table.sort(dangling_nums)
    local parts = {}
    for _, n in ipairs(dangling_nums) do
      parts[#parts + 1] = "[" .. n .. "]"
    end
    ref_diag_warn("以下引用编号在参考文献中无对应条目，已保留纯文本: " .. table.concat(parts, " "))
  end
  return out
end

-- ============================================================
-- 3. 裸图片段落升格为 Figure
--    管线为 mermaid/图表生成的替换图片 alt 留空（避免与作者手写
--    图注重复产生双图注），而 pandoc 的 implicit_figures 只对非空
--    alt 生效，空 alt 图片因此留在普通段落——docx 落 Body Text
--    样式（继承首行缩进，图片被缩进两字符）。这里在 AST 层补齐
--    同一语义：段落中仅有一张图片时升格为 Figure——docx 落
--    Figure 样式（无缩进），html 落 <figure> 语义元素，输出一致。
-- ============================================================

local function promote_bare_image_paragraphs(blocks)
  for i, block in ipairs(blocks) do
    if block.t == "Para" and #block.content == 1 and block.content[1].t == "Image" then
      blocks[i] = pandoc.Figure({ pandoc.Plain({ block.content[1] }) }, {})
    end
  end
end

-- ============================================================
-- 6. 水平线（---）→ 分页符（docx / typst）
--    pandoc docx writer 把 HorizontalRule 渲染为 VML 横线形状
--    （<w:pict><v:rect o:hr="t"/>，Word 中呈现为图片对象），
--    与正式文档里 --- 的本意（强制分页）不符。
--    html 输出保留 <hr>（网页无分页概念）。
--    去重：模板 Heading 1 样式自带段前分页（# 章 自动另起新页），
--    紧跟一级标题的 --- 再转分页符会叠加成空白页——前瞻下一个
--    块，是一级标题或文档末尾时丢弃。故在 Pandoc handler 内处理
--    （需要块上下文做前瞻），不用元素级 handler。
--    typst 输出的一级标题不带自动分页，PDF 中 --- 无条件转分页。
-- ============================================================

local function is_hr_pagebreak_redundant(doc, i)
  local nxt = doc.blocks[i + 1]
  if nxt == nil then
    return true -- 文档末尾的分页符只会产生空白尾页
  end
  return nxt.t == "Header" and nxt.level == 1
end

function Pandoc(doc)
  local out = {}
  local buffer = {}
  local inside = false

  local function flush_buffer()
    if #buffer == 0 then
      return
    end
    local html = table.concat(buffer, "\n")
    buffer = {}
    local blocks = wrap_html_table(html)
    for _, block in ipairs(blocks) do
      -- HTML 重组出的原生表在 Pandoc handler 内部生成，不会重走
      -- Table handler 遍历，列宽归一化必须在这里显式补上——
      -- 否则重组表永远得不到按内容分配的列宽
      if block.t == "Table" then
        block = assign_column_widths(block) or block
      end
      out[#out + 1] = block
    end
  end

  for i, block in ipairs(doc.blocks) do
    if block.t == "HorizontalRule" then
      -- 见第 6 节：docx/PDF 转分页符；紧跟一级标题或文末时去重丢弃；html 原样
      if FORMAT:match("docx") then
        if not is_hr_pagebreak_redundant(doc, i) then
          out[#out + 1] = pandoc.RawBlock("openxml", '<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
        end
      elseif FORMAT:match("typst") then
        out[#out + 1] = pandoc.RawBlock("typst", "#pagebreak()")
      else
        out[#out + 1] = block
      end
    elseif block.t == "RawBlock" and block.format == "html" then
      if is_table_start(block.text) then
        inside = true
        buffer[#buffer + 1] = block.text
        -- 单块自包含表格（<table>...</table> 同块）立即闭缓冲，
        -- 否则其后所有内容都会被卷进 HTML 重解析管线直到 EOF
        if is_table_end(block.text) then
          inside = false
          flush_buffer()
        end
      elseif inside then
        buffer[#buffer + 1] = block.text
        if is_table_end(block.text) then
          inside = false
          flush_buffer()
        end
      else
        out[#out + 1] = block
      end
    else
      if inside then
        buffer[#buffer + 1] = pandoc.write(
          pandoc.Pandoc({ math_to_raw_literal(block) }), "html")
      else
        out[#out + 1] = block
      end
    end
  end

  if inside then
    -- <table> 有开无闭：缓冲一直吞到文档尾，其后所有内容都被卷入
    -- HTML 重解析——留下诊断让作者能定位漏写的 </table>
    ref_diag_warn("检测到未闭合的 <table> 标签，其后内容被卷入表格重解析，请检查标签配对")
    flush_buffer()
  end

  local final_blocks = apply_citation_pass(out, doc.meta)
  promote_bare_image_paragraphs(final_blocks)
  return pandoc.Pandoc(final_blocks, doc.meta)
end

-- ============================================================
-- 4. RawInline 行内 HTML -> 原生 Inline
--    markdown reader 把行内 HTML 标签保留为 RawInline，
--    docx writer 会直接丢弃（<br> 换行丢失、<b> 变普通文本），
--    这里转换成原生元素，Word / HTML 两种输出都能正确生效。
-- ============================================================

-- 成对标签 -> 对应的原生构造器
local PAIR_TAGS = {
  b = 'Strong', strong = 'Strong',
  i = 'Emph', em = 'Emph',
  u = 'Underline',
  sub = 'Subscript', sup = 'Superscript',
  del = 'Strikeout', s = 'Strikeout', strike = 'Strikeout',
}

local function build_pair_frame(tag)
  return { build = function(content) return pandoc[PAIR_TAGS[tag]](content) end, buf = pandoc.List() }
end

local function build_link_frame(open_text)
  -- href 三种形态：双引号 / 单引号 / 无引号裸值（<a href=foo>）
  local href = open_text:match('href%s*=%s*"([^"]*)"')
      or open_text:match("href%s*=%s*'([^']+)'")
      or open_text:match("href%s*=%s*([^%s>'\"]+)")
      or ''
  return {
    build = function(content) return pandoc.Link(content, href, '') end,
    buf = pandoc.List(),
  }
end

local function build_image(text)
  local src = text:match('src%s*=%s*"([^"]*)"') or text:match("src%s*=%s*'([^']+)'")
  if not src then
    return nil
  end
  local alt = text:match('alt%s*=%s*"([^"]*)"') or text:match("alt%s*=%s*'([^']+)'") or ''
  return pandoc.Image(pandoc.List({ pandoc.Str(alt) }), src, '')
end

function Inlines(inlines)
  local has_raw = false
  for _, el in ipairs(inlines) do
    if el.t == 'RawInline' and el.format == 'html' then
      has_raw = true
      break
    end
  end
  if not has_raw then
    return nil
  end

  local out = pandoc.List()
  local stack = pandoc.List()

  local function target_list()
    if #stack > 0 then
      return stack[#stack].buf
    end
    return out
  end

  local function pop_frame()
    local frame = stack:remove(#stack)
    if frame then
      target_list():insert(frame.build(frame.buf))
    end
    return frame
  end

  for _, el in ipairs(inlines) do
    if el.t == 'RawInline' and el.format == 'html' then
      local text = el.text
      local open_tag = text:match('^<([%a]+)[^>]*>$')
      local close_tag = text:match('^</([%a]+)[^>]*>$')

      if close_tag then
        if close_tag == 'a' or PAIR_TAGS[close_tag] then
          if not pop_frame() then
            -- 游离闭合标签（无对应的开标签）：按未知标签保留，
            -- 与其走同一条路——静默丢弃会让文字凭空消失
            target_list():insert(el)
          end
        else
          target_list():insert(el)
        end
      elseif open_tag then
        if open_tag == 'br' then
          target_list():insert(pandoc.LineBreak())
        elseif open_tag == 'img' then
          local image = build_image(text)
          if image then
            target_list():insert(image)
          end
        elseif open_tag == 'a' then
          stack:insert(build_link_frame(text))
        elseif PAIR_TAGS[open_tag] then
          stack:insert(build_pair_frame(open_tag))
        else
          target_list():insert(el)
        end
      else
        target_list():insert(el)
      end
    else
      target_list():insert(el)
    end
  end

  -- 未闭合的标签：内容按包裹处理，避免丢字
  while #stack > 0 do
    pop_frame()
  end

  return out
end

-- ============================================================
-- 5. 表格列宽归一化
--    pandoc 只对"源码行宽超过 --columns"的管道表计算相对列宽；
--    较窄的表列宽为 null，docx writer 随之输出 autofit 窄表。
--    这里在 AST 层给无列宽的表按单元格内容长度分配比例列宽
--    （和为 1），writer 即确定性输出 tblW=100% 的固定布局表格，
--    全部撑满正文区宽度。
-- ============================================================

local stringify = pandoc.utils.stringify

-- 估算字符串显示宽度：CJK/全角按 2、其余按 1。直接用字节长度（#s）
-- 会让中文列权重虚高约 1.5 倍，中英混排表的列宽分配系统性失真
local function display_width(text)
  local width = 0
  local ok = pcall(function()
    for _, codepoint in utf8.codes(text) do
      local wide = (codepoint >= 0x1100 and codepoint <= 0x115F) -- 谚文
          or (codepoint >= 0x2E80 and codepoint <= 0x9FFF) -- CJK 部首/符号/表意/假名
          or (codepoint >= 0xAC00 and codepoint <= 0xD7A3) -- 谚文音节
          or (codepoint >= 0xF900 and codepoint <= 0xFAFF) -- CJK 兼容表意
          or (codepoint >= 0xFF00 and codepoint <= 0xFF60) -- 全角形式
          or (codepoint >= 0xFFE0 and codepoint <= 0xFFE6) -- 全角符号
      width = width + (wide and 2 or 1)
    end
  end)
  if not ok then
    return #text -- 非法 UTF-8：退回字节长度
  end
  return width
end

-- 空内容列的保底权重（约 3 个汉字宽），避免被长文本列挤成 0 宽
local MIN_WEIGHT = 6
-- 归一化后每列保底占比：短"标签列"在长文本列旁不至于只剩一条缝
local MIN_FRACTION = 0.08

-- 该 pandoc 版本经 Lua 读取 colspec 的 width 恒为 nil（读侧序列化缺陷，
-- 连 pandoc 自己计算的宽度也读不到），只能借 native 序列化文本获取：
-- 返回（是否含 ColWidthDefault，pandoc 计算出的宽度值列表）。
-- 只取前 ncols 个 token：native 序列化中本表 colspecs 位于表头/表体
-- 之前，嵌套表的 ColWidth 出现在单元格内容之后，全局 gmatch 会被
-- 嵌套表污染（计数错位 → 显式比例误判为需重排）
local function table_width_info(tbl)
  local native = pandoc.write(pandoc.Pandoc({ tbl }), "native")
  local ncols = #tbl.colspecs
  local widths = {}
  local has_default = false
  local pos = 1
  for _ = 1, ncols do
    local w_start, w_end, value = native:find("ColWidth%s+([%d.eE+-]+)", pos)
    local d_start, d_end = native:find("ColWidthDefault", pos)
    if d_start and (not w_start or d_start < w_start) then
      has_default = true
      pos = d_end + 1
    elseif w_start then
      widths[#widths + 1] = tonumber(value)
      pos = w_end + 1
    else
      break
    end
  end
  return has_default, widths
end

-- 赋值给文件头前置声明的同名局部：HTML 重组表（Pandoc handler 内）
-- 与 Table handler 都要调用本函数
function assign_column_widths(tbl)
  local ncols = #tbl.colspecs
  if ncols == 0 then
    return nil
  end
  local has_default, pandoc_widths = table_width_info(tbl)
  if not has_default and #pandoc_widths == ncols then
    -- pandoc 已计算列宽且各列相等：这是 |---|---|---| 等长破折号分隔行
    -- 的默认写法（源码行超宽时 pandoc 按破折号比例算宽），不携带作者的
    -- 比例意图，均分会把长文本列挤窄——仍按内容重新分配；
    -- 宽度不全相等才是显式比例（不等宽分隔行），尊重不动。
    -- 含 ColWidthDefault（无列宽/混合形态，writer 按全默认处理）同样走内容分配
    for _, w in ipairs(pandoc_widths) do
      if math.abs(w - pandoc_widths[1]) > 1e-6 then
        return nil
      end
    end
  end

  local weights = {}
  for i = 1, ncols do
    weights[i] = 0
  end
  for _, row in ipairs(table_rows(tbl)) do
    for col_i, cell in ipairs(row.cells) do
      if col_i > ncols then
        break
      end
      local span = math.max(tonumber(cell.col_span) or 1, 1)
      -- 跨列单元格的内容均摊到其覆盖的列（按显示宽度，非字节长度）
      local len = display_width(stringify(cell)) / span
      for k = 0, span - 1 do
        local idx = col_i + k
        if idx > ncols then
          break
        end
        weights[idx] = math.max(weights[idx], len)
      end
    end
  end

  local total = 0
  for i = 1, ncols do
    weights[i] = math.max(weights[i], MIN_WEIGHT)
    total = total + weights[i]
  end

  local fracs = {}
  for i = 1, ncols do
    fracs[i] = weights[i] / total
  end

  -- 保底列宽一轮再分配：过窄列提到 MIN_FRACTION，宽列按比例让出
  local deficit, flexible = 0, 0
  for _, f in ipairs(fracs) do
    if f < MIN_FRACTION then
      deficit = deficit + MIN_FRACTION - f
    else
      flexible = flexible + f
    end
  end
  if deficit > 0 and flexible > deficit then
    for i, f in ipairs(fracs) do
      if f < MIN_FRACTION then
        fracs[i] = MIN_FRACTION
      else
        fracs[i] = f * (flexible - deficit) / flexible
      end
    end
  end

  -- 消除舍入漂移：再次归一化到和为 1。
  -- colspecs 访问器返回副本，且写侧序列化为列表形态（[1]=align 字符串，
  -- [2]=宽度数字），必须整体构造后赋值回去，原地改/命名字段均不生效
  local sum = 0
  for _, f in ipairs(fracs) do
    sum = sum + f
  end
  local specs = {}
  for i = 1, ncols do
    specs[i] = { tbl.colspecs[i][1] or "AlignDefault", fracs[i] / sum }
  end
  tbl.colspecs = specs
  return tbl
end

function Table(tbl)
  return assign_column_widths(tbl) or tbl
end
