-- Fit figures to the text width and size table columns from their text.
-- A 100% width does not enlarge a tiny image on pandoc 3.1; images that are
-- wider than the page shrink to the A4 text width from reference.docx.

function Image(img)
  img.attributes["width"] = "100%"
  img.attributes["height"] = nil
  return img
end

local function cell_len(cell)
  local blocks = cell.contents or cell
  local text = pandoc.utils.stringify(blocks)
  local length = utf8.len(text) or #text
  if length < 1 then
    length = 1
  end
  return length
end

local function consider(weights, row, boost)
  if not row or not row.cells then
    return
  end
  for index, cell in ipairs(row.cells) do
    local weight = cell_len(cell) * (boost or 1)
    weights[index] = math.max(weights[index] or 1, weight)
  end
end

function Table(tbl)
  local columns = #tbl.colspecs
  if columns < 1 then
    return nil
  end

  local weights = {}
  for index = 1, columns do
    weights[index] = 1
  end

  if tbl.head and tbl.head.rows then
    for _, row in ipairs(tbl.head.rows) do
      consider(weights, row, 1.35)
    end
  end

  if tbl.bodies then
    for _, body in ipairs(tbl.bodies) do
      for _, row in ipairs(body.head or {}) do
        consider(weights, row, 1.2)
      end
      for _, row in ipairs(body.body or {}) do
        consider(weights, row, 1)
      end
    end
  end

  local total = 0
  for index = 1, columns do
    if weights[index] > 48 then
      weights[index] = 48
    end
    total = total + weights[index]
  end

  local shares = {}
  local sum = 0
  local min_share = math.min(0.08, 1 / columns)
  for index = 1, columns do
    local share = weights[index] / total
    if share < min_share then
      share = min_share
    end
    shares[index] = share
    sum = sum + share
  end

  for index = 1, columns do
    local align = tbl.colspecs[index][1]
    tbl.colspecs[index] = { align, shares[index] / sum }
  end

  return tbl
end
