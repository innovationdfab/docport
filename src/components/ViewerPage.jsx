import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
} from 'react'

import ExcelGrid from './ExcelGrid.jsx'
import CustomSheetViews from './CustomSheetViews.jsx'
import {
  workbook,
  sheetOrder,
} from '../data/workbook.js'

const API_BASE_URL = ''
const customViews = {
  Invoice: [
    'fixedInvoiceView',
    '.invoice-sheet-inner',
  ],

  'Dhl Express SLI': [
    'fixedDhlView',
    '.dhl-sheet-inner',
  ],

  'FORM SDF': [
    'fixedSdfView',
    '.sdf-sheet-inner',
  ],

  EVD: [
    'fixedEvdView',
    '.evd-sheet-inner',
  ],

  'AlumSteel Derivatives declarati': [
    'fixedAlumView',
    '.alum-sheet-inner',
  ],

  'Scomet Declaration': [
    'fixedScometView',
    '.scomet-sheet-inner',
  ],
}


export default function ViewerPage({
  onBack,
  showToast,
  pdfFile,
  analysisResult,
  documentData,
  onFieldChange,
  onLineItemChange,
}) {
  const [
    activeSheet,
    setActiveSheet,
  ] = useState(
    sheetOrder[0]
  )

  const [
    zoom,
    setZoom,
  ] = useState(1)

  const [
    selectedCell,
    setSelectedCell,
  ] = useState({
    address: 'A1',
    value: '',
  })


  // =========================================================
  // REAL EXCEL GENERATION STATES
  // =========================================================

  const [
    generatingExcel,
    setGeneratingExcel,
  ] = useState(false)

  const [
    generatedExcelBlob,
    setGeneratedExcelBlob,
  ] = useState(null)

  const [
    generatedFileName,
    setGeneratedFileName,
  ] = useState(
    'Export_Documents.xlsx'
  )

  const [
    generatedDataSignature,
    setGeneratedDataSignature,
  ] = useState('')


  const viewportRef =
    useRef(null)

  const canvasRef =
    useRef(null)

  const customHostRef =
    useRef(null)

  const customBaseScaleRef =
    useRef(1)


  const sheet =
    useMemo(
      () =>
        workbook[
          activeSheet
        ],

      [
        activeSheet,
      ]
    )


  const customConfig =
    customViews[
      activeSheet
    ]


  // =========================================================
  // CURRENT UPLOADED PDF NAME
  // =========================================================

  const viewerFileName =
    pdfFile?.name ||
    analysisResult?.filename ||
    'Export Documents'


  // =========================================================
  // SIGNATURE OF CURRENT DATA
  // =========================================================

  const currentDataSignature =
    useMemo(
      () => {
        try {
          return JSON.stringify(
            documentData || {}
          )
        } catch {
          return ''
        }
      },

      [
        documentData,
      ]
    )


  const excelIsCurrent =
    generatedExcelBlob &&
    generatedDataSignature ===
      currentDataSignature


  // =========================================================
  // CUSTOM SHEET SCALE
  // =========================================================

  const applyCustomScale =
    useCallback(
      (
        requestedZoom =
          zoom
      ) => {
        if (
          !customConfig ||
          !viewportRef.current ||
          !customHostRef.current
        ) {
          return false
        }


        const [
          viewId,
          innerSelector,
        ] =
          customConfig


        const outer =
          customHostRef.current
            .querySelector(
              `#${viewId}`
            )


        const inner =
          outer?.querySelector(
            innerSelector
          )


        if (
          !outer ||
          !inner
        ) {
          return false
        }


        inner.style.transform =
          'none'

        outer.style.height =
          'auto'


        const available =
          Math.max(
            320,

            viewportRef.current
              .clientWidth -
              14
          )


        const natural =
          inner.scrollWidth ||
          1180


        const baseScale =
          Math.min(
            1,
            available /
              natural
          )


        customBaseScaleRef.current =
          baseScale


        const scale =
          Math.max(
            0.45,

            Math.min(
              1.7,

              baseScale *
                requestedZoom
            )
          )


        inner.style.transform =
          `scale(${scale})`


        outer.style.height =
          `${
            inner.scrollHeight *
            scale
          }px`


        return true
      },

      [
        customConfig,
        zoom,
      ]
    )


  // =========================================================
  // ACTIVE SHEET DISPLAY
  // =========================================================

  useLayoutEffect(
    () => {
      if (
        !customHostRef.current
      ) {
        return
      }


      Object.entries(
        customViews
      ).forEach(
        ([
          sheetName,
          [viewId],
        ]) => {
          const element =
            customHostRef.current
              .querySelector(
                `#${viewId}`
              )


          if (element) {
            element.style.display =
              sheetName ===
              activeSheet
                ? 'block'
                : 'none'
          }
        }
      )


      if (
        canvasRef.current
      ) {
        canvasRef.current
          .style.visibility =
          customConfig
            ? 'hidden'
            : 'visible'
      }


      if (
        customConfig
      ) {
        applyCustomScale(
          1
        )
      }


      viewportRef.current
        ?.scrollTo(
          0,
          0
        )
    },

    [
      activeSheet,
      customConfig,
      applyCustomScale,
    ]
  )


  // =========================================================
  // WINDOW RESIZE
  // =========================================================

  useEffect(
    () => {
      const onResize =
        () => {
          if (
            customConfig
          ) {
            applyCustomScale(
              zoom
            )
          }
        }


      window.addEventListener(
        'resize',
        onResize
      )


      return () => {
        window
          .removeEventListener(
            'resize',
            onResize
          )
      }
    },

    [
      customConfig,
      applyCustomScale,
      zoom,
    ]
  )


  // =========================================================
  // CHANGE SHEET
  // =========================================================

  const changeSheet =
    useCallback(
      (name) => {
        setActiveSheet(
          name
        )

        setZoom(1)

        setSelectedCell({
          address: 'A1',
          value: '',
        })
      },

      []
    )


  // =========================================================
  // ZOOM
  // =========================================================

  const zoomBy =
    useCallback(
      (delta) => {
        setZoom(
          (current) => {
            const next =
              Math.max(
                0.45,

                Math.min(
                  1.7,

                  current +
                    delta
                )
              )


            window
              .requestAnimationFrame(
                () => {
                  if (
                    customConfig
                  ) {
                    applyCustomScale(
                      next
                    )
                  }
                }
              )


            showToast(
              `Zoom ${Math.round(
                next * 100
              )}%`
            )


            return next
          }
        )
      },

      [
        applyCustomScale,
        customConfig,
        showToast,
      ]
    )


  // =========================================================
  // FIT WIDTH
  // =========================================================

  const fitWidth =
    useCallback(
      () => {
        if (
          customConfig
        ) {
          setZoom(1)


          window
            .requestAnimationFrame(
              () =>
                applyCustomScale(
                  1
                )
            )


          showToast(
            'Fit to width'
          )

          return
        }


        const viewport =
          viewportRef.current

        const canvas =
          canvasRef.current


        if (
          !viewport ||
          !canvas
        ) {
          return
        }


        const naturalWidth =
          Number.parseFloat(
            canvas.style.width
          ) ||
          canvas.scrollWidth


        const next =
          Math.min(
            1,

            (
              viewport
                .clientWidth -
              12
            ) /
              naturalWidth
          )


        setZoom(next)


        showToast(
          'Fit to width'
        )
      },

      [
        applyCustomScale,
        customConfig,
        showToast,
      ]
    )


  // =========================================================
  // GET FILE NAME FROM BACKEND
  // =========================================================

  const getDownloadFileName =
    useCallback(
      (response) => {
        const disposition =
          response.headers.get(
            'content-disposition'
          )


        if (
          !disposition
        ) {
          const invoiceNumber =
            documentData
              ?.extracted
              ?.invoice
              ?.number


          return invoiceNumber
            ? `Export_Documents_${invoiceNumber}.xlsx`
            : 'Export_Documents.xlsx'
        }


        const utfMatch =
          disposition.match(
            /filename\*=UTF-8''([^;]+)/i
          )


        if (
          utfMatch?.[1]
        ) {
          return decodeURIComponent(
            utfMatch[1]
              .replace(
                /["']/g,
                ''
              )
          )
        }


        const normalMatch =
          disposition.match(
            /filename="?([^"]+)"?/i
          )


        if (
          normalMatch?.[1]
        ) {
          return normalMatch[1]
            .replace(
              /[";]/g,
              ''
            )
            .trim()
        }


        return (
          'Export_Documents.xlsx'
        )
      },

      [
        documentData,
      ]
    )


  // =========================================================
  // GENERATE REAL EXCEL
  // =========================================================

  const handleGenerateExcel =
    useCallback(
      async () => {
        if (
          !documentData
        ) {
          showToast(
            'No analyzed invoice data available.'
          )

          return
        }


        if (
          generatingExcel
        ) {
          return
        }


        try {
          setGeneratingExcel(
            true
          )


          showToast(
            'Generating Excel...'
          )


          const response =
            await fetch(
              `${API_BASE_URL}/api/documents/generate`,

              {
                method:
                  'POST',

                headers: {
                  'Content-Type':
                    'application/json',
                },

                body:
                  JSON.stringify(
                    documentData
                  ),
              }
            )


          if (
            !response.ok
          ) {
            let errorMessage =
              'Excel generation failed.'


            const contentType =
              response.headers.get(
                'content-type'
              ) || ''


            try {
              if (
                contentType.includes(
                  'application/json'
                )
              ) {
                const errorData =
                  await response.json()


                errorMessage =
                  errorData
                    ?.detail ||

                  errorData
                    ?.message ||

                  errorMessage
              } else {
                const errorText =
                  await response.text()


                if (
                  errorText
                ) {
                  errorMessage =
                    errorText
                }
              }
            } catch {
              // Keep normal message.
            }


            throw new Error(
              errorMessage
            )
          }


          const blob =
            await response.blob()


          if (
            !blob ||
            blob.size === 0
          ) {
            throw new Error(
              'Backend returned an empty Excel file.'
            )
          }


          const fileName =
            getDownloadFileName(
              response
            )


          setGeneratedExcelBlob(
            blob
          )


          setGeneratedFileName(
            fileName
          )


          setGeneratedDataSignature(
            currentDataSignature
          )


          console.log(
            'Generated Excel:',

            {
              fileName,
              size:
                blob.size,
              type:
                blob.type,
            }
          )


          showToast(
            'Excel generated successfully.'
          )
        } catch (
          error
        ) {
          console.error(
            'Generate Excel error:',

            error
          )


          setGeneratedExcelBlob(
            null
          )


          setGeneratedDataSignature(
            ''
          )


          showToast(
            error?.message ||
              'Unable to generate Excel.'
          )
        } finally {
          setGeneratingExcel(
            false
          )
        }
      },

      [
        documentData,
        generatingExcel,
        showToast,
        getDownloadFileName,
        currentDataSignature,
      ]
    )


  // =========================================================
  // DOWNLOAD CURRENT REVIEWED EXCEL
  // =========================================================

  const handleDownloadExcel =
    useCallback(
      async () => {
        if (
          !documentData
        ) {
          showToast(
            'No analyzed invoice data available.'
          )

          return
        }


        if (
          generatingExcel
        ) {
          return
        }


        try {
          setGeneratingExcel(
            true
          )


          showToast(
            'Preparing Excel download...'
          )


          const response =
            await fetch(
              `${API_BASE_URL}/api/documents/generate`,

              {
                method:
                  'POST',

                headers: {
                  'Content-Type':
                    'application/json',
                },

                body:
                  JSON.stringify(
                    documentData
                  ),
              }
            )


          if (
            !response.ok
          ) {
            let errorMessage =
              'Excel generation failed.'


            const contentType =
              response.headers.get(
                'content-type'
              ) || ''


            try {
              if (
                contentType.includes(
                  'application/json'
                )
              ) {
                const errorData =
                  await response.json()


                errorMessage =
                  errorData
                    ?.detail ||

                  errorData
                    ?.message ||

                  errorMessage
              } else {
                const errorText =
                  await response.text()


                if (
                  errorText
                ) {
                  errorMessage =
                    errorText
                }
              }
            } catch {
              // Keep normal message.
            }


            throw new Error(
              errorMessage
            )
          }


          const blob =
            await response.blob()


          if (
            !blob ||
            blob.size === 0
          ) {
            throw new Error(
              'Backend returned an empty Excel file.'
            )
          }


          const fileName =
            getDownloadFileName(
              response
            )


          const url =
            window.URL
              .createObjectURL(
                blob
              )


          const link =
            document
              .createElement(
                'a'
              )


          link.href =
            url

          link.download =
            fileName ||
            'Export_Documents.xlsx'


          document.body
            .appendChild(
              link
            )


          link.click()

          link.remove()


          window.setTimeout(
            () => {
              window.URL
                .revokeObjectURL(
                  url
                )
            },

            1000
          )


          showToast(
            'Excel downloaded.'
          )
        } catch (
          error
        ) {
          console.error(
            'Download Excel error:',

            error
          )


          showToast(
            error?.message ||
              'Unable to download Excel.'
          )
        } finally {
          setGeneratingExcel(
            false
          )
        }
      },

      [
        documentData,
        generatingExcel,
        showToast,
        getDownloadFileName,
      ]
    )


  // =========================================================
  // PAGE
  // =========================================================

  return (
    <div
      id="viewerPage"
      className="page viewerPage active"
    >
      <header
        className="viewerHeader"
      >
        <div
          className="viewerLeft"
        >
          <button
            className="btn"
            type="button"
            onClick={onBack}
          >
            ← Back
          </button>


          <div>
            <div
              className="fileTitle"
            >
              {viewerFileName} — Generated
            </div>


            <div
              className="fileSub"
            >
              All six original sheets • no shortened wording
            </div>
          </div>
        </div>


        <div
          className="viewerRight"
        >
          <button
            className="btn"
            type="button"
            onClick={
              fitWidth
            }
          >
            Fit Width
          </button>


          <button
            className="btn"
            type="button"
            onClick={() =>
              zoomBy(-0.1)
            }
            aria-label="Zoom out"
          >
            − Zoom
          </button>


          <button
            className="btn"
            type="button"
            onClick={() =>
              zoomBy(0.1)
            }
            aria-label="Zoom in"
          >
            + Zoom
          </button>


          <button
            className="btn download"
            type="button"
            onClick={
              handleDownloadExcel
            }
            disabled={
              generatingExcel
            }
            style={{
              opacity:
                generatingExcel
                  ? 0.7
                  : 1,

              cursor:
                generatingExcel
                  ? 'wait'
                  : 'pointer',
            }}
          >
            {
              generatingExcel
                ? 'Preparing Download...'
                : '⬇ Download Excel'
            }
          </button>
        </div>
      </header>


      <div
        className="excelShell"
      >
        {/*
          Formula bar removed intentionally.
          No A1 / fx / formula strip.
        */}


        <div
          id="viewport"
          ref={viewportRef}
          className="viewport"
        >
          <ExcelGrid
            sheetName={
              activeSheet
            }
            sheet={sheet}
            zoom={zoom}
            canvasRef={
              canvasRef
            }
            onCellSelect={(
              address,
              value
            ) =>
              setSelectedCell({
                address,
                value,
              })
            }
          />


          <CustomSheetViews
            ref={
              customHostRef
            }
            documentData={
              documentData
            }
            analysisResult={
              analysisResult
            }
            onFieldChange={
              onFieldChange
            }
            onLineItemChange={
              onLineItemChange
            }
          />
        </div>


        <div
          id="tabs"
          className="tabs"
          role="tablist"
          aria-label="Workbook sheets"
        >
          {
            sheetOrder.map(
              (name) => (
                <button
                  type="button"
                  key={name}
                  className={`tab${
                    name ===
                    activeSheet
                      ? ' active'
                      : ''
                  }`}
                  onClick={() =>
                    changeSheet(
                      name
                    )
                  }
                  role="tab"
                  aria-selected={
                    name ===
                    activeSheet
                  }
                >
                  {name}
                </button>
              )
            )
          }
        </div>
      </div>


      
    </div>
  )
}