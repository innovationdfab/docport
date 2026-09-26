import { useMemo } from 'react'

function colName(index) {
  let n = index + 1
  let result = ''
  while (n) {
    const r = (n - 1) % 26
    result = String.fromCharCode(65 + r) + result
    n = Math.floor((n - 1) / 26)
  }
  return result
}

function normFill(fill) {
  return !fill || fill === '#FFFFFF' ? 'transparent' : fill
}

function borderCss(border) {
  if (!border || !border.width) return 'none'
  return `${border.width}px ${border.style || 'solid'} ${border.color || '#000'}`
}

export default function ExcelGrid({ sheetName, sheet, zoom, onCellSelect, canvasRef }) {
  const layout = useMemo(() => {
    const x = [0]
    const y = [0]
    sheet.widths.forEach((value) => x.push(x[x.length - 1] + value))
    sheet.heights.forEach((value) => y.push(y[y.length - 1] + value))
    return { x, y, gridW: x[x.length - 1], gridH: y[y.length - 1] }
  }, [sheet])

  const anchors = useMemo(() => {
    const map = new Map()
    sheet.boxes.forEach((box) => map.set(`${box.r}:${box.c}`, box))
    return map
  }, [sheet])

  const rowHeaderWidth = 46
  const columnHeaderHeight = 25

  return (
    <div
      id="canvas"
      ref={canvasRef}
      className="canvas"
      style={{
        width: `${rowHeaderWidth + layout.gridW}px`,
        height: `${columnHeaderHeight + layout.gridH}px`,
        transform: `scale(${zoom})`,
      }}
    >
      <div className="corner" />

      {sheet.widths.map((width, column) => (
        <div
          className="colHead"
          key={`col-${column}`}
          style={{ left: `${rowHeaderWidth + layout.x[column]}px`, width: `${width}px` }}
        >
          {colName(column)}
        </div>
      ))}

      {sheet.heights.map((height, row) => (
        <div
          className="rowHead"
          key={`row-${row}`}
          style={{ top: `${columnHeaderHeight + layout.y[row]}px`, height: `${height}px` }}
        >
          {row + 1}
        </div>
      ))}

      <div className="grid" style={{ width: `${layout.gridW}px`, height: `${layout.gridH}px` }}>
        {layout.x.map((left, index) => (
          <div className="gridV" key={`gv-${index}`} style={{ left: `${left}px`, height: `${layout.gridH}px` }} />
        ))}
        {layout.y.map((top, index) => (
          <div className="gridH" key={`gh-${index}`} style={{ top: `${top}px`, width: `${layout.gridW}px` }} />
        ))}

        {sheet.boxes.map((box, index) => {
          const style = box.s || {}
          const left = layout.x[box.c]
          const top = layout.y[box.r]
          const right = layout.x[Math.min(sheet.widths.length, box.c + box.cs)]
          const bottom = layout.y[Math.min(sheet.rows, box.r + box.rs)]
          return (
            <div
              className="cellBg"
              key={`bg-${index}`}
              style={{
                left: `${left}px`,
                top: `${top}px`,
                width: `${right - left}px`,
                height: `${bottom - top}px`,
                background: normFill(style.fill),
                borderLeft: borderCss(style.bL),
                borderRight: borderCss(style.bR),
                borderTop: borderCss(style.bT),
                borderBottom: borderCss(style.bB),
              }}
            />
          )
        })}

        {sheet.boxes.map((box, index) => {
          if (box.v === null || box.v === undefined || String(box.v) === '') return null

          const style = box.s || {}
          const left = layout.x[box.c]
          const top = layout.y[box.r]
          let endColumn = Math.min(sheet.widths.length, box.c + box.cs)

          if (!style.wrap && box.cs === 1) {
            let scan = box.c + 1
            while (scan < sheet.widths.length) {
              const next = anchors.get(`${box.r}:${scan}`)
              if (next && next.v !== null && next.v !== undefined && String(next.v) !== '') break
              scan += 1
            }
            endColumn = Math.max(endColumn, scan)
          }

          const right = layout.x[endColumn]
          const bottom = layout.y[Math.min(sheet.rows, box.r + box.rs)]
          const useItalic = Boolean(style.italic) && sheetName !== 'Invoice'
          const classes = [
            'cellText',
            style.wrap ? 'wrap' : 'nowrap',
            style.bold ? 'bold' : '',
            style.underline ? 'underline' : '',
            useItalic ? 'italic' : '',
            style.h === 'center' ? 'center' : style.h === 'right' ? 'right' : '',
          ].filter(Boolean).join(' ')

          let family = style.fontFamily || 'Arial'
          if (family === 'Calibri, Arial') family = 'Calibri, Arial, sans-serif'
          const address = `${colName(box.c)}${box.r + 1}`
          const value = String(box.v)

          return (
            <button
              type="button"
              className={classes}
              key={`text-${index}`}
              onClick={() => onCellSelect(address, value)}
              style={{
                left: `${left}px`,
                top: `${top}px`,
                width: `${right - left}px`,
                height: `${bottom - top}px`,
                fontFamily: family,
                fontSize: `${style.fontSize || 10}pt`,
                color: style.fontColor || '#000',
                fontStyle: useItalic ? 'italic' : 'normal',
                justifyContent: style.h === 'center' ? 'center' : style.h === 'right' ? 'flex-end' : 'flex-start',
                alignItems: style.v === 'top' ? 'flex-start' : style.v === 'center' ? 'center' : 'flex-end',
                border: 0,
                cursor: 'default',
              }}
              aria-label={`${address}: ${value}`}
            >
              <span className="txt">{value}</span>
            </button>
          )
        })}
      </div>
    </div>
  )
}
