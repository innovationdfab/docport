import {
  useCallback,
  useEffect,
  useRef,
  useState,
} from 'react'

import HomePage from './components/HomePage'
import ViewerPage from './components/ViewerPage'

const API_BASE_URL = 'http://127.0.0.1:8000'

export default function App() {
  const [page, setPage] = useState('home')
  const [pdfFile, setPdfFile] = useState(null)

  const [analyzed, setAnalyzed] = useState(false)
  const [analysisResult, setAnalysisResult] = useState(null)

  const [manualEdits, setManualEdits] = useState({})

  const [progress, setProgress] = useState(0)
  const [status, setStatus] = useState('')
  const [toastMessage, setToastMessage] = useState('')

  const timerRef = useRef([])
  const toastTimerRef = useRef(null)

  // =========================================================
  // FINAL DOCUMENT DATA
  // PDF DATA + MANUAL CORRECTIONS
  // =========================================================

  const finalDocumentData = analysisResult
    ? {
        ...analysisResult,

        extracted: {
          ...analysisResult.extracted,

          invoice: {
            ...analysisResult.extracted?.invoice,
            ...(manualEdits.invoice || {}),
          },

          exporter: {
            ...analysisResult.extracted?.exporter,
            ...(manualEdits.exporter || {}),
          },

          importer: {
            ...analysisResult.extracted?.importer,
            ...(manualEdits.importer || {}),
          },

          consignee: {
            ...analysisResult.extracted?.consignee,
            ...(manualEdits.consignee || {}),
          },

          sold_to: {
            ...analysisResult.extracted?.sold_to,
            ...(manualEdits.sold_to || {}),
          },

          shipment: {
            ...analysisResult.extracted?.shipment,
            ...(manualEdits.shipment || {}),
          },

          product: {
            ...analysisResult.extracted?.product,
            ...(manualEdits.product || {}),
          },

          company: {
            ...analysisResult.extracted?.company,
            ...(manualEdits.company || {}),
          },

          bank: {
            ...analysisResult.extracted?.bank,
            ...(manualEdits.bank || {}),
          },

          ui_overrides: {
            ...analysisResult.extracted?.ui_overrides,
            ...(manualEdits.ui_overrides || {}),
          },
        },
      }
    : null

  // =========================================================
  // TOAST
  // =========================================================

  const showToast = useCallback((message) => {
    setToastMessage(message)

    if (toastTimerRef.current) {
      window.clearTimeout(toastTimerRef.current)
    }

    toastTimerRef.current = window.setTimeout(() => {
      setToastMessage('')
    }, 2200)
  }, [])

  // =========================================================
  // CLEAR TIMERS
  // =========================================================

  const clearTimers = useCallback(() => {
    timerRef.current.forEach((timer) => {
      window.clearTimeout(timer)
    })

    timerRef.current = []
  }, [])

  // =========================================================
  // NORMAL MANUAL FIELD EDIT
  // =========================================================

  const handleFieldChange = useCallback(
    (section, field, value) => {
      setManualEdits((previous) => ({
        ...previous,

        [section]: {
          ...(previous[section] || {}),
          [field]: value,
        },
      }))
    },
    []
  )

  // =========================================================
  // PRODUCT ROW EDIT
  // =========================================================

  const handleLineItemChange = useCallback(
    (index, field, value) => {
      if (!analysisResult) {
        return
      }

      const currentItems =
        finalDocumentData?.extracted?.product?.line_items || []

      const updatedItems = currentItems.map(
        (item, itemIndex) => {
          if (itemIndex !== index) {
            return item
          }

          return {
            ...item,
            [field]: value,
          }
        }
      )

      setManualEdits((previous) => ({
        ...previous,

        product: {
          ...(previous.product || {}),
          line_items: updatedItems,
        },
      }))
    },
    [analysisResult, finalDocumentData]
  )

  // =========================================================
  // FILE CHANGE
  // =========================================================

  const handleFileChange = useCallback(
    (file) => {
      clearTimers()

      setPdfFile(file)
      setAnalyzed(false)
      setAnalysisResult(null)
      setManualEdits({})
      setProgress(0)
      setStatus('')

      if (!file) {
        return
      }

      const isPdf =
        file.type === 'application/pdf' ||
        file.name.toLowerCase().endsWith('.pdf')

      if (!isPdf) {
        setPdfFile(null)
        showToast('Please select a PDF file.')
      }
    },
    [clearTimers, showToast]
  )

  // =========================================================
  // ANALYZE PDF
  // =========================================================

  const handleAnalyze = async () => {
    if (!pdfFile) {
      showToast('Upload the Commercial Invoice PDF first.')
      return
    }

    try {
      clearTimers()

      setAnalyzed(false)
      setAnalysisResult(null)
      setManualEdits({})

      setProgress(15)
      setStatus('Reading PDF...')

      const formData = new FormData()

      formData.append('file', pdfFile)

      setProgress(40)

      setStatus(
        'Extracting invoice and product details...'
      )

      const response = await fetch(
        `${API_BASE_URL}/api/invoices/extract`,
        {
          method: 'POST',
          body: formData,
        }
      )

      let result

      try {
        result = await response.json()
      } catch {
        throw new Error(
          'Backend returned an invalid response.'
        )
      }

      if (!response.ok) {
        throw new Error(
          result?.detail ||
            result?.message ||
            'PDF analysis failed.'
        )
      }

      setProgress(80)

      setStatus(
        'Matching values with all 6 Excel sheets...'
      )

      setAnalysisResult(result)

      sessionStorage.setItem(
        'exportflow_analysis',
        JSON.stringify(result)
      )

      const productCount =
        result?.extracted?.product?.line_items?.length || 0

      setProgress(100)
      setAnalyzed(true)

      setStatus(
        `Analysis complete. ${productCount} ${
          productCount === 1
            ? 'product'
            : 'products'
        } found. Ready to generate Excel.`
      )

      console.log(
        'PDF ANALYSIS RESULT:',
        result
      )

      console.log(
        'EXTRACTED PRODUCTS:',
        result?.extracted?.product?.line_items || []
      )

      console.log(
        'MISSING FIELDS:',
        result?.missing_fields || []
      )

      showToast(
        `PDF analyzed successfully. ${productCount} ${
          productCount === 1
            ? 'product'
            : 'products'
        } found.`
      )
    } catch (error) {
      console.error(
        'Analyze PDF error:',
        error
      )

      setProgress(0)
      setAnalyzed(false)
      setAnalysisResult(null)
      setManualEdits({})

      const message =
        error?.message ||
        'Unable to analyze the PDF.'

      setStatus(message)
      showToast(message)
    }
  }

  // =========================================================
  // HOME GENERATE
  // Opens preview only
  // =========================================================

  const handleGenerate = useCallback(() => {
    if (!pdfFile) {
      showToast('Upload the PDF first.')
      return
    }

    if (!analyzed) {
      showToast('Analyze the PDF first.')
      return
    }

    if (!analysisResult) {
      showToast(
        'PDF analysis data is not available.'
      )

      return
    }

    setPage('viewer')

    window.scrollTo(0, 0)
  }, [
    pdfFile,
    analyzed,
    analysisResult,
    showToast,
  ])

  // =========================================================
  // BACK
  // =========================================================

  const handleBack = useCallback(() => {
    setPage('home')

    window.scrollTo(0, 0)
  }, [])

  // =========================================================
  // CLEANUP
  // =========================================================

  useEffect(() => {
    return () => {
      clearTimers()

      if (toastTimerRef.current) {
        window.clearTimeout(
          toastTimerRef.current
        )
      }
    }
  }, [clearTimers])

  // =========================================================
  // UI
  // =========================================================

  return (
    <>
      {page === 'home' ? (
        <HomePage
          pdfFile={pdfFile}
          progress={progress}
          status={status}
          onFileChange={handleFileChange}
          onAnalyze={handleAnalyze}
          onGenerate={handleGenerate}
        />
      ) : (
        <ViewerPage
          onBack={handleBack}
          onGoBack={handleBack}
          goBack={handleBack}

          pdfFile={pdfFile}

          analysisResult={analysisResult}

          documentData={finalDocumentData}

          manualEdits={manualEdits}

          onFieldChange={handleFieldChange}

          onLineItemChange={handleLineItemChange}

          showToast={showToast}
        />
      )}

      <div
        id="toast"
        className="toast"
        style={{
          display: toastMessage
            ? 'block'
            : 'none',
        }}
      >
        {toastMessage}
      </div>
    </>
  )
}