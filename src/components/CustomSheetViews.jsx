import {
  forwardRef,
  useCallback,
  useLayoutEffect,
  useRef,
} from 'react'

import { customViewsHtml } from '../data/customViews.js'


// ============================================================
// HELPERS
// ============================================================

const isMissing = (value) =>
  value === null ||
  value === undefined ||
  String(value).trim() === ''


const cleanNumber = (value) => {
  if (isMissing(value)) return ''

  return String(value)
    .replace(/,/g, '')
    .replace(/[^\d.-]/g, '')
}


const numberValue = (value) => {
  const cleaned = cleanNumber(value)

  if (!cleaned) return null

  const parsed = Number(cleaned)

  return Number.isFinite(parsed)
    ? parsed
    : null
}


const formatMoney = (
  value,
  symbol = ''
) => {
  const number =
    numberValue(value)

  if (number === null) {
    return ''
  }

  return `${symbol}${number.toFixed(2)}`
}


const parseDateParts = (value) => {
  if (isMissing(value)) {
    return null
  }

  const text =
    String(value).trim()

  let match =
    text.match(
      /^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/
    )

  if (match) {
    return {
      day:
        match[1].padStart(
          2,
          '0'
        ),

      month:
        match[2].padStart(
          2,
          '0'
        ),

      year:
        match[3],
    }
  }

  match =
    text.match(
      /^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/
    )

  if (match) {
    return {
      day:
        match[3].padStart(
          2,
          '0'
        ),

      month:
        match[2].padStart(
          2,
          '0'
        ),

      year:
        match[1],
    }
  }

  return null
}


const slashDate = (value) => {
  const parts =
    parseDateParts(value)

  if (!parts) {
    return isMissing(value)
      ? ''
      : String(value)
  }

  return (
    `${parts.day}/` +
    `${parts.month}/` +
    `${parts.year}`
  )
}


const dashDate = (value) => {
  const parts =
    parseDateParts(value)

  if (!parts) {
    return isMissing(value)
      ? ''
      : String(value)
  }

  return (
    `${parts.day}-` +
    `${parts.month}-` +
    `${parts.year}`
  )
}


const getCurrencySymbol = (
  shipment,
  lineItems
) => {
  const firstSymbol =
    lineItems?.find(
      (item) =>
        !isMissing(
          item?.currency_symbol
        )
    )?.currency_symbol

  if (firstSymbol) {
    return firstSymbol
  }

  const currency =
    shipment?.currency

  if (
    !currency ||
    currency === 'USD'
  ) {
    return '$'
  }

  if (currency === 'INR') {
    return '₹'
  }

  if (currency === 'EUR') {
    return '€'
  }

  if (currency === 'GBP') {
    return '£'
  }

  return String(currency)
}


const getTotalAmount = (
  shipment,
  lineItems
) => {
  const shipmentAmount =
    numberValue(
      shipment?.total_amount
    )

  if (
    shipmentAmount !== null
  ) {
    return shipmentAmount
  }

  if (
    !Array.isArray(lineItems)
  ) {
    return null
  }

  const values =
    lineItems
      .map(
        (item) =>
          numberValue(
            item?.position_price
          )
      )
      .filter(
        (value) =>
          value !== null
      )

  if (!values.length) {
    return null
  }

  return values.reduce(
    (sum, value) =>
      sum + value,
    0
  )
}


const getTotalNetWeight = (
  shipment,
  lineItems
) => {
  const shipmentWeight =
    numberValue(
      shipment?.net_weight_kg
    )

  if (
    shipmentWeight !== null
  ) {
    return shipmentWeight
  }

  const values =
    (lineItems || [])
      .map(
        (item) =>
          numberValue(
            item?.net_weight_kg
          )
      )
      .filter(
        (value) =>
          value !== null
      )

  if (!values.length) {
    return null
  }

  return values.reduce(
    (sum, value) =>
      sum + value,
    0
  )
}


const productDescription = (
  item
) => {
  if (!item) {
    return ''
  }

  if (
    !isMissing(
      item.description
    )
  ) {
    return String(
      item.description
    )
  }

  const parts = []

  if (
    !isMissing(
      item.part_number
    )
  ) {
    parts.push(
      `Part #${item.part_number}`
    )
  }

  if (
    !isMissing(
      item.material
    )
  ) {
    parts.push(
      `Material: ${item.material}`
    )
  }

  return parts.join('\n')
}


const allProductDescription = (
  lineItems
) =>
  (lineItems || [])
    .map(
      (item, index) => {
        const description =
          productDescription(
            item
          )

        if (!description) {
          return ''
        }

        return (
          lineItems.length > 1
            ? `Product ${index + 1}: ${description}`
            : description
        )
      }
    )
    .filter(Boolean)
    .join('\n')


const findRow = (
  table,
  text
) => {
  if (!table) {
    return null
  }

  return Array.from(
    table.rows || []
  ).find(
    (row) =>
      row.textContent
        ?.toLowerCase()
        .includes(
          text.toLowerCase()
        )
  )
}


// ============================================================
// COMPONENT
// ============================================================

const CustomSheetViews =
  forwardRef(
    function CustomSheetViews(
      {
        documentData,
        analysisResult,
        onFieldChange,
        onLineItemChange,
      },
      forwardedRef
    ) {
      const hostRef =
        useRef(null)


      const setHostRef =
        useCallback(
          (node) => {
            hostRef.current =
              node

            if (
              typeof forwardedRef ===
              'function'
            ) {
              forwardedRef(node)
            } else if (
              forwardedRef
            ) {
              forwardedRef.current =
                node
            }
          },
          [forwardedRef]
        )


      useLayoutEffect(() => {
        const root =
          hostRef.current

        if (!root) {
          return
        }


        const extracted =
          documentData?.extracted ||
          analysisResult?.extracted ||
          {}


        const invoice =
          extracted.invoice ||
          {}

        const consignee =
          extracted.consignee ||
          {}

        const shipment =
          extracted.shipment ||
          {}

        const company =
          extracted.company ||
          {}

        const bank =
          extracted.bank ||
          {}

        const product =
          extracted.product ||
          {}


        const lineItems =
          Array.isArray(
            product.line_items
          )
            ? product.line_items
            : []


        const totalAmount =
          getTotalAmount(
            shipment,
            lineItems
          )


        const totalNetWeight =
          getTotalNetWeight(
            shipment,
            lineItems
          )


        const exchangeRate =
          numberValue(
            shipment.exchange_rate
          )


        const amountInr =
          totalAmount !== null &&
          exchangeRate !== null
            ? totalAmount *
              exchangeRate
            : null


        const currencySymbol =
          getCurrencySymbol(
            shipment,
            lineItems
          )


        const consigneeName =
          consignee.name ||
          consignee.contact ||
          ''


        const consigneeAddress =
          consignee.raw_section ||
          consignee.address ||
          ''


        const documentDate =
          invoice.date ||
          shipment.shipping_bill_date ||
          ''


        const invoiceDate =
          documentDate


        const shippingBillDate =
          shipment.shipping_bill_date ||
          documentDate


        const countryOrigin =
          lineItems.find(
            (item) =>
              !isMissing(
                item.country_of_origin
              )
          )?.country_of_origin ||
          'INDIA'


        // ====================================================
        // EDITABLE CELL HANDLER
        // ====================================================

        const bindEditable = (
          element,
          {
            value,
            onSave,
            display,
            parse,
            required = true,
            multiline = false,
          }
        ) => {
          if (!element) {
            return
          }


          const rawValue =
            isMissing(value)
              ? ''
              : value


          const shown =
            display
              ? display(rawValue)
              : String(rawValue)


          element.classList.add(
            'dynamic-editable'
          )


          element.contentEditable =
            'true'

          element.spellcheck =
            false

          element.dataset.dynamic =
            'true'


          element.classList.toggle(
            'missing-cell',
            required &&
              isMissing(rawValue)
          )


          element.title =
            required &&
            isMissing(rawValue)
              ? 'Missing value — click here and enter it'
              : 'Click to edit'


          if (
            element.dataset.editing !==
            'true'
          ) {
            element.textContent =
              shown
          }


          element.onfocus = () => {
            element.dataset.editing =
              'true'
          }


          element.oninput = () => {
            const typed =
              element.innerText.trim()

            element.dataset.pendingValue =
              typed

            element.classList.toggle(
              'missing-cell',
              required &&
                typed === ''
            )
          }


          element.onblur = () => {
            const typed =
              (
                element.dataset.pendingValue ??
                element.innerText
              ).trim()


            const finalValue =
              parse
                ? parse(typed)
                : typed


            element.dataset.editing =
              'false'

            element.dataset.pendingValue =
              ''


            onSave?.(
              finalValue
            )


            const savedShown =
              display
                ? display(
                    finalValue
                  )
                : String(
                    finalValue ?? ''
                  )


            element.textContent =
              savedShown


            element.classList.toggle(
              'missing-cell',
              required &&
                isMissing(
                  finalValue
                )
            )


            element.title =
              required &&
              isMissing(
                finalValue
              )
                ? 'Missing value — click here and enter it'
                : 'Click to edit'
          }


          element.onkeydown =
            (event) => {
              if (
                !multiline &&
                event.key ===
                  'Enter'
              ) {
                event.preventDefault()

                element.blur()
              }
            }
        }


        const bindField = (
          element,
          value,
          section,
          field,
          options = {}
        ) => {
          bindEditable(
            element,
            {
              ...options,
              value,

              onSave:
                (
                  nextValue
                ) => {
                  onFieldChange?.(
                    section,
                    field,
                    nextValue
                  )
                },
            }
          )
        }


        const bindLineField = (
          element,
          value,
          index,
          field,
          options = {}
        ) => {
          bindEditable(
            element,
            {
              ...options,
              value,

              onSave:
                (
                  nextValue
                ) => {
                  onLineItemChange?.(
                    index,
                    field,
                    nextValue
                  )
                },
            }
          )
        }


        const setNormalValue = (
          element,
          value
        ) => {
          if (!element) {
            return
          }

          element.textContent =
            isMissing(value)
              ? ''
              : String(value)
        }


        const setDerivedValue = (
          element,
          value,
          {
            missingWhenEmpty =
              false,
          } = {}
        ) => {
          if (!element) {
            return
          }

          element.textContent =
            isMissing(value)
              ? ''
              : String(value)


          element.classList.toggle(
            'missing-cell',
            missingWhenEmpty &&
              isMissing(value)
          )


          if (
            missingWhenEmpty &&
            isMissing(value)
          ) {
            element.title =
              'Complete the related missing field'
          }
        }


        const makeDynamicSpan = (
          cell,
          prefix,
          value,
          section,
          field,
          {
            suffix = '',
            display,
            parse,
          } = {}
        ) => {
          if (!cell) {
            return
          }


          cell.replaceChildren()


          cell.appendChild(
            document.createTextNode(
              prefix
            )
          )


          const span =
            document.createElement(
              'span'
            )


          span.className =
            'inline-dynamic-field'


          cell.appendChild(
            span
          )


          if (suffix) {
            cell.appendChild(
              document.createTextNode(
                suffix
              )
            )
          }


          bindField(
            span,
            value,
            section,
            field,
            {
              display,
              parse,
            }
          )
        }


        // ====================================================
        // 1. INVOICE SHEET
        // ====================================================

        const invoiceTable =
          root.querySelector(
            '#fixedInvoiceView .invoice-main'
          )


        if (invoiceTable) {
          const rows =
            invoiceTable.rows


          // CONSIGNEE

          const topRow =
            rows[1]


          const consigneeTable =
            topRow?.cells?.[0]
              ?.querySelector(
                '.mini'
              )


          if (consigneeTable) {
            const cRows =
              consigneeTable.rows


            const consigneeNameCell =
              cRows?.[5]
                ?.cells?.[0]
                ?.querySelector(
                  '.blue-fill'
                )


            bindField(
              consigneeNameCell,
              consigneeName,
              'consignee',
              'name'
            )


            const addressCell =
              cRows?.[6]
                ?.cells?.[0]


            bindField(
              addressCell,
              consigneeAddress,
              'consignee',
              'raw_section',
              {
                multiline:
                  true,
              }
            )


            if (
              addressCell?.style
            ) {
              addressCell.style.whiteSpace =
                'pre-line'
            }


            const contactCell =
              cRows?.[8]
                ?.cells?.[0]
                ?.querySelector(
                  '.blue-fill'
                )


            bindField(
              contactCell,
              consigneeName,
              'consignee',
              'name'
            )
          }


          // INVOICE NUMBER / DATE / PO

          const invoiceInfoTable =
            topRow?.cells?.[2]
              ?.querySelector(
                '.mini'
              )


          if (
            invoiceInfoTable
          ) {
            const infoRows =
              invoiceInfoTable.rows


            bindField(
              infoRows?.[0]
                ?.cells?.[1],

              invoiceDate,

              'invoice',
              'date',

              {
                display:
                  (value) =>
                    value
                      ? `Date-${slashDate(
                          value
                        )}`
                      : '',

                parse:
                  (value) =>
                    value.replace(
                      /^Date[-:\s]*/i,
                      ''
                    ),
              }
            )


            bindField(
              infoRows?.[1]
                ?.cells?.[0],

              invoice.number,

              'invoice',
              'number'
            )


            bindField(
              infoRows?.[2]
                ?.cells?.[1],

              invoiceDate,

              'invoice',
              'date',

              {
                display:
                  (value) =>
                    value
                      ? `Date-${slashDate(
                          value
                        )}`
                      : '',

                parse:
                  (value) =>
                    value.replace(
                      /^Date[-:\s]*/i,
                      ''
                    ),
              }
            )


            bindField(
              infoRows?.[3]
                ?.cells?.[0],

              invoice.po_number,

              'invoice',
              'po_number'
            )


            setNormalValue(
              infoRows?.[5]
                ?.cells?.[0],

              countryOrigin
            )


            setNormalValue(
              infoRows?.[7]
                ?.cells?.[0],

              countryOrigin
            )
          }


          // EXCHANGE RATE

          const paymentTable =
            rows?.[2]
              ?.cells?.[1]
              ?.querySelector(
                '.mini'
              )


          if (paymentTable) {
            bindField(
              paymentTable
                .rows?.[0]
                ?.cells?.[1],

              shipment.exchange_rate,

              'shipment',
              'exchange_rate',

              {
                parse:
                  cleanNumber,
              }
            )
          }


          // CARRIER

          const carriageTable =
            rows?.[3]
              ?.cells?.[0]
              ?.querySelector(
                '.mini'
              )


          if (carriageTable) {
            setNormalValue(
              carriageTable
                .rows?.[3]
                ?.cells?.[0],

              invoice.carrier ||
                'DHL'
            )
          }


          // LINE ITEMS

          const lineTable =
            root.querySelector(
              '#fixedInvoiceView .line-items'
            )


          if (lineTable) {
            while (
              lineTable.rows.length >
              1
            ) {
              lineTable.deleteRow(
                1
              )
            }


            lineItems.forEach(
              (item, index) => {
                const row =
                  lineTable.insertRow()


                const createCell =
                  (
                    className = ''
                  ) => {
                    const cell =
                      row.insertCell()

                    cell.className =
                      className

                    return cell
                  }


                setNormalValue(
                  createCell(
                    'center'
                  ),
                  index + 1
                )


                setNormalValue(
                  createCell(
                    'center'
                  ),
                  item.india_hsn ||
                    '76169990'
                )


                setNormalValue(
                  createCell(
                    'center'
                  ),
                  item.hs_code
                )


                const descCell =
                  createCell(
                    'blue-box center'
                  )


                descCell.style.whiteSpace =
                  'pre-line'


                bindLineField(
                  descCell,

                  productDescription(
                    item
                  ),

                  index,

                  'description',

                  {
                    multiline:
                      true,
                  }
                )


                setNormalValue(
                  createCell(
                    'center'
                  ),
                  item.quantity
                )


                const unitPriceCell =
                  createCell(
                    'blue-fill'
                  )


                bindLineField(
                  unitPriceCell,

                  item.unit_price,

                  index,

                  'unit_price',

                  {
                    display:
                      (value) =>
                        formatMoney(
                          value,
                          item.currency_symbol ||
                            currencySymbol
                        ),

                    parse:
                      cleanNumber,
                  }
                )


                const taxableCell =
                  createCell(
                    'blue-fill bold center'
                  )


                bindLineField(
                  taxableCell,

                  item.position_price,

                  index,

                  'position_price',

                  {
                    display:
                      (value) =>
                        formatMoney(
                          value
                        ),

                    parse:
                      cleanNumber,
                  }
                )


                setNormalValue(
                  createCell(
                    'center bold'
                  ),
                  '0%'
                )


                setNormalValue(
                  createCell(
                    'center bold'
                  ),
                  '0'
                )
              }
            )
          }


          // TOTALS / BANK / PACKING

          const totalsRow =
            rows?.[5]


          if (totalsRow) {
            const amountTable =
              totalsRow.cells?.[0]
                ?.querySelector(
                  '.mini'
                )


            const packingTable =
              totalsRow.cells?.[1]
                ?.querySelector(
                  '.mini'
                )


            const finalTotalTable =
              totalsRow.cells?.[2]
                ?.querySelector(
                  '.mini'
                )


            if (amountTable) {
              setDerivedValue(
                amountTable
                  .rows?.[0]
                  ?.cells?.[1],

                amountInr === null
                  ? ''
                  : amountInr.toFixed(
                      2
                    ),

                {
                  missingWhenEmpty:
                    true,
                }
              )


              bindField(
                amountTable
                  .rows?.[1]
                  ?.cells?.[1],

                totalAmount,

                'shipment',
                'total_amount',

                {
                  display:
                    (value) =>
                      formatMoney(
                        value
                      ),

                  parse:
                    cleanNumber,
                }
              )


              if (
                !isMissing(
                  company.iec
                )
              ) {
                setNormalValue(
                  amountTable
                    .rows?.[2]
                    ?.cells?.[1],

                  company.iec
                )
              }


              if (
                !isMissing(
                  company.gstin
                )
              ) {
                setNormalValue(
                  amountTable
                    .rows?.[3]
                    ?.cells?.[1],

                  company.gstin
                )
              }


              if (
                !isMissing(
                  bank.code
                )
              ) {
                setNormalValue(
                  amountTable
                    .rows?.[4]
                    ?.cells?.[1],

                  bank.code
                )
              }


              if (
                !isMissing(
                  bank.account_number
                )
              ) {
                setNormalValue(
                  amountTable
                    .rows?.[5]
                    ?.cells?.[1],

                  bank.account_number
                )
              }
            }


            if (packingTable) {
              bindField(
                packingTable
                  .rows?.[2]
                  ?.cells?.[0]
                  ?.querySelector(
                    '.blue-fill'
                  ),

                shipment.package_dimensions,

                'shipment',
                'package_dimensions'
              )


              bindField(
                packingTable
                  .rows?.[3]
                  ?.cells?.[0]
                  ?.querySelector(
                    '.blue-fill'
                  ),

                shipment.package_weight_kg,

                'shipment',
                'package_weight_kg',

                {
                  display:
                    (value) =>
                      value
                        ? `${value} kg`
                        : '',

                  parse:
                    cleanNumber,
                }
              )


              bindField(
                packingTable
                  .rows?.[4]
                  ?.cells?.[0]
                  ?.querySelector(
                    '.blue-fill'
                  ),

                totalNetWeight,

                'shipment',
                'net_weight_kg',

                {
                  display:
                    (value) =>
                      value
                        ? `${value} kg`
                        : '',

                  parse:
                    cleanNumber,
                }
              )
            }


            if (
              finalTotalTable
            ) {
              bindField(
                finalTotalTable
                  .rows?.[0]
                  ?.cells?.[2],

                totalAmount,

                'shipment',
                'total_amount',

                {
                  display:
                    (value) =>
                      formatMoney(
                        value
                      ),

                  parse:
                    cleanNumber,
                }
              )


              bindField(
                finalTotalTable
                  .rows?.[2]
                  ?.cells?.[2],

                totalAmount,

                'shipment',
                'total_amount',

                {
                  display:
                    (value) =>
                      formatMoney(
                        value
                      ),

                  parse:
                    cleanNumber,
                }
              )
            }
          }
        }


        // ====================================================
        // 2. DHL EXPRESS SLI
        // ====================================================

        const dhlTable =
          root.querySelector(
            '#fixedDhlView > .dhl-sheet-inner > table'
          )


        if (dhlTable) {
          const invoiceRow =
            findRow(
              dhlTable,
              'Invoice No.:'
            )


          bindField(
            invoiceRow
              ?.cells?.[3],

            invoice.number,

            'invoice',
            'number'
          )


          const consigneeRow =
            findRow(
              dhlTable,
              'Consignee Name:'
            )


          bindField(
            consigneeRow
              ?.cells?.[1],

            consigneeName,

            'consignee',
            'name'
          )


          bindField(
            consigneeRow
              ?.cells?.[3],

            invoiceDate,

            'invoice',
            'date',

            {
              display:
                dashDate,
            }
          )


          const awbRow =
            findRow(
              dhlTable,
              'DHL AIR WAYBILL NUMBER'
            )


          bindField(
            awbRow
              ?.cells?.[1],

            shipment.awb_number,

            'shipment',
            'awb_number'
          )


          const taxRow =
            findRow(
              dhlTable,
              'Export Against Payment'
            )


          if (
            taxRow?.cells?.[1]
          ) {
            const taxCell =
              taxRow.cells[1]


            taxCell.innerHTML =
              `TAXABLE AMOUNT -${
                amountInr === null
                  ? ''
                  : amountInr.toFixed(
                      2
                    )
              }<br/>` +
              'IGST RATE - 0<br/>' +
              'IGST AMOUNT - 0<br/>' +
              'GST Compensation Cess - 0'


            taxCell.classList.toggle(
              'missing-cell',
              amountInr === null
            )


            if (
              amountInr === null
            ) {
              taxCell.title =
                'Enter the missing exchange rate'
            }
          }


          const fobRow =
            findRow(
              dhlTable,
              'FOB VALUE'
            )


          bindField(
            fobRow
              ?.cells?.[2],

            totalAmount,

            'shipment',
            'total_amount',

            {
              display:
                (value) =>
                  value
                    ? `${formatMoney(
                        value
                      )} USD`
                    : '',

              parse:
                cleanNumber,
            }
          )


          const netRow =
            findRow(
              dhlTable,
              'NET WT.'
            )


          bindField(
            netRow
              ?.cells?.[2],

            totalNetWeight,

            'shipment',
            'net_weight_kg',

            {
              display:
                (value) =>
                  value
                    ? `${value} kg`
                    : '',

              parse:
                cleanNumber,
            }
          )


          const grossRow =
            findRow(
              dhlTable,
              'GROSS WT.'
            )


          bindField(
            grossRow
              ?.cells?.[2],

            shipment.package_weight_kg,

            'shipment',
            'package_weight_kg',

            {
              display:
                (value) =>
                  value
                    ? `${value} kg`
                    : '',

              parse:
                cleanNumber,
            }
          )
        }


        // ====================================================
        // 3. FORM SDF
        // ====================================================

        const sdfInfo =
          root.querySelector(
            '#fixedSdfView .sdf-info'
          )


        if (sdfInfo) {
          const row =
            sdfInfo.rows?.[0]


          bindField(
            row?.cells?.[1],

            shipment.shipping_bill_number,

            'shipment',
            'shipping_bill_number'
          )


          bindField(
            row?.cells?.[3],

            shippingBillDate,

            'shipment',
            'shipping_bill_date',

            {
              display:
                slashDate,
            }
          )
        }


        const sdfDeclaration =
          root.querySelector(
            '#fixedSdfView .blue-soft.sdf-text'
          )


        if (sdfDeclaration) {
          sdfDeclaration.replaceChildren()


          sdfDeclaration.append(
            document.createTextNode(
              'We here by declare that we are the   SELLER / CONSIGNOR   of the goods in Respect of which this declaration made and that particulars given to shipping Bill No: '
            )
          )


          const billSpan =
            document.createElement(
              'span'
            )


          sdfDeclaration.append(
            billSpan
          )


          bindField(
            billSpan,

            shipment.shipping_bill_number,

            'shipment',
            'shipping_bill_number'
          )


          sdfDeclaration.append(
            document.createTextNode(
              ' Date '
            )
          )


          const dateSpan =
            document.createElement(
              'span'
            )


          sdfDeclaration.append(
            dateSpan
          )


          bindField(
            dateSpan,

            shippingBillDate,

            'shipment',
            'shipping_bill_date',

            {
              display:
                slashDate,
            }
          )


          sdfDeclaration.append(
            document.createTextNode(
              ' are true and that'
            )
          )
        }


        // ====================================================
        // 4. EVD
        // ====================================================

        const evdTable =
          root.querySelector(
            '#fixedEvdView .evd-main'
          )


        if (evdTable) {
          const shippingRow =
            findRow(
              evdTable,
              '1. Shipping Bill No.'
            )


          if (
            shippingRow
              ?.cells?.[0]
          ) {
            const cell =
              shippingRow.cells[0]


            cell.replaceChildren()


            cell.append(
              document.createTextNode(
                '1. Shipping Bill No. '
              )
            )


            const billSpan =
              document.createElement(
                'span'
              )


            cell.append(
              billSpan
            )


            bindField(
              billSpan,

              shipment.shipping_bill_number,

              'shipment',
              'shipping_bill_number'
            )


            cell.append(
              document.createTextNode(
                '     & Date:- '
              )
            )


            const dateSpan =
              document.createElement(
                'span'
              )


            cell.append(
              dateSpan
            )


            bindField(
              dateSpan,

              shippingBillDate,

              'shipment',
              'shipping_bill_date',

              {
                display:
                  slashDate,
              }
            )
          }


          const invoiceRow =
            findRow(
              evdTable,
              '2. Invoice No.'
            )


          if (invoiceRow) {
            const firstCell =
              invoiceRow
                .cells?.[0]


            if (firstCell) {
              firstCell.replaceChildren()


              firstCell.append(
                document.createTextNode(
                  '2. Invoice No. & Date   '
                )
              )


              const numberSpan =
                document.createElement(
                  'span'
                )


              firstCell.append(
                numberSpan
              )


              bindField(
                numberSpan,

                invoice.number,

                'invoice',
                'number'
              )


              firstCell.append(
                document.createTextNode(
                  ' & '
                )
              )


              const dateSpan =
                document.createElement(
                  'span'
                )


              firstCell.append(
                dateSpan
              )


              bindField(
                dateSpan,

                invoiceDate,

                'invoice',
                'date',

                {
                  display:
                    slashDate,
                }
              )
            }


            makeDynamicSpan(
              invoiceRow
                .cells?.[1],

              'Date: ',

              invoiceDate,

              'invoice',
              'date',

              {
                display:
                  slashDate,
              }
            )
          }


          const previousRow =
            findRow(
              evdTable,
              'Shipping Bill No:'
            )


          if (
            previousRow &&
            previousRow !==
              shippingRow
          ) {
            const cell =
              previousRow
                .cells?.[0]


            if (cell) {
              cell.replaceChildren()


              cell.append(
                document.createTextNode(
                  'Shipping Bill No: '
                )
              )


              const billSpan =
                document.createElement(
                  'span'
                )


              cell.append(
                billSpan
              )


              bindField(
                billSpan,

                shipment.shipping_bill_number,

                'shipment',
                'shipping_bill_number'
              )


              cell.append(
                document.createTextNode(
                  ' and date: '
                )
              )


              const dateSpan =
                document.createElement(
                  'span'
                )


              cell.append(
                dateSpan
              )


              bindField(
                dateSpan,

                shippingBillDate,

                'shipment',
                'shipping_bill_date',

                {
                  display:
                    slashDate,
                }
              )
            }
          }


          const dateRows =
            Array.from(
              evdTable.rows
            ).filter(
              (row) =>
                row.textContent
                  ?.trim()
                  .startsWith(
                    'Date:'
                  )
            )


          const finalDateRow =
            dateRows[
              dateRows.length -
                1
            ]


          if (
            finalDateRow
              ?.cells?.[0]
          ) {
            makeDynamicSpan(
              finalDateRow
                .cells[0],

              'Date: ',

              shippingBillDate,

              'shipment',
              'shipping_bill_date',

              {
                display:
                  slashDate,
              }
            )
          }
        }


        // ====================================================
        // 5. ALUMINIUM / STEEL DECLARATION
        // ====================================================

        const alumTable =
          root.querySelector(
            '#fixedAlumView .alum-main'
          )


        if (alumTable) {
          const productRow =
            alumTable.querySelector(
              'td.blue.small'
            )


          if (productRow) {
            productRow.style.whiteSpace =
              'pre-line'


            setNormalValue(
              productRow,

              allProductDescription(
                lineItems
              )
            )


            productRow.classList.toggle(
              'missing-cell',

              !allProductDescription(
                lineItems
              )
            )
          }


          const weightRow =
            findRow(
              alumTable,
              'Full weight of the product'
            )


          if (
            weightRow
              ?.cells?.[0] &&

            weightRow.cells[0]
              .classList.contains(
                'red'
              )
          ) {
            const cell =
              weightRow.cells[0]


            cell.innerHTML =
              `3) Full weight of the product&nbsp;&nbsp;${
                totalNetWeight ??
                ''
              }_kg<br/>` +
              `Aluminum content weight&nbsp;${
                totalNetWeight ??
                ''
              }_kg`


            cell.classList.toggle(
              'missing-cell',

              totalNetWeight ===
                null
            )
          }


          const valueRow =
            findRow(
              alumTable,
              'Total Value of the product'
            )


          if (
            valueRow
              ?.cells?.[0] &&

            valueRow.cells[0]
              .classList.contains(
                'blue'
              )
          ) {
            const cell =
              valueRow.cells[0]


            cell.innerHTML =
              `Total Value of the product&nbsp;&nbsp;${
                totalAmount === null
                  ? ''
                  : totalAmount.toFixed(
                      2
                    )
              } usd<br/>` +
              `Value of the Aluminum content&nbsp;&nbsp;${
                totalAmount === null
                  ? ''
                  : totalAmount.toFixed(
                      2
                    )
              } usd`


            cell.classList.toggle(
              'missing-cell',

              totalAmount ===
                null
            )
          }


          const completedRow =
            findRow(
              alumTable,
              'Completed by'
            )


          if (
            completedRow
              ?.cells?.[0]
          ) {
            makeDynamicSpan(
              completedRow
                .cells[0],

              'Completed by : ',

              invoiceDate,

              'invoice',
              'date',

              {
                display:
                  dashDate,
              }
            )
          }


          const titleRow =
            findRow(
              alumTable,
              'Title'
            )


          if (
            titleRow
              ?.cells?.[0]
          ) {
            titleRow
              .cells[0]
              .replaceChildren()


            titleRow
              .cells[0]
              .append(
                document.createTextNode(
                  'Title        : '
                )
              )


            const titleSpan =
              document.createElement(
                'span'
              )


            titleRow
              .cells[0]
              .append(
                titleSpan
              )


            const firstProduct =
              lineItems?.[0]


            bindLineField(
              titleSpan,

              firstProduct
                ?.description ||

                firstProduct
                  ?.part_number ||

                '',

              0,

              firstProduct
                ?.description
                ? 'description'
                : 'part_number',

              {
                multiline:
                  true,
              }
            )
          }


          const allDateRows =
            Array.from(
              alumTable.rows
            ).filter(
              (row) =>
                row.textContent
                  ?.trim()
                  .startsWith(
                    'Date'
                  )
            )


          const bottomDateRow =
            allDateRows[
              allDateRows.length -
                1
            ]


          if (
            bottomDateRow
              ?.cells?.[0]
          ) {
            makeDynamicSpan(
              bottomDateRow
                .cells[0],

              'Date          : ',

              invoiceDate,

              'invoice',
              'date',

              {
                display:
                  dashDate,
              }
            )
          }
        }


        // ====================================================
        // 6. SCOMET DECLARATION
        // ====================================================

        const scometTable =
          root.querySelector(
            '#fixedScometView .scomet-main'
          )


        if (scometTable) {
          const invoiceRow =
            findRow(
              scometTable,
              'Invoice No:'
            )


          if (invoiceRow) {
            makeDynamicSpan(
              invoiceRow
                .cells?.[0],

              'Invoice No: - ',

              invoice.number,

              'invoice',
              'number'
            )


            makeDynamicSpan(
              invoiceRow
                .cells?.[1],

              'Date: ',

              invoiceDate,

              'invoice',
              'date',

              {
                display:
                  dashDate,
              }
            )
          }


          const productRow =
            findRow(
              scometTable,
              'Product description'
            )


          if (
            productRow
              ?.cells?.[0]
          ) {
            const cell =
              productRow.cells[0]


            cell.style.whiteSpace =
              'pre-line'


            const description =
              allProductDescription(
                lineItems
              )


            setNormalValue(
              cell,

              description
                ? `Product description:- ${description}`
                : ''
            )


            cell.classList.toggle(
              'missing-cell',

              !description
            )
          }


          const hsnRow =
            findRow(
              scometTable,
              'HSN code'
            )


          if (
            hsnRow
              ?.cells?.[0]
          ) {
            const hsnText =
              lineItems
                .map(
                  (item) =>
                    item.hs_code
                )
                .filter(Boolean)
                .join(', ')


            setNormalValue(
              hsnRow.cells[0],

              hsnText
                ? `HSN code: - ${hsnText}`
                : 'HSN code: -'
            )
          }


          // ==================================================
          // FORCE COMPLETE END-USE SENTENCE
          // ==================================================

          const scometRows =
            Array.from(
              scometTable.rows ||
                []
            )


          const endUseRow =
            scometRows[
              scometRows.length -
                1
            ]


          const endUseCell =
            endUseRow
              ?.cells?.[0]


          if (endUseCell) {
            endUseCell.textContent =
              'End Use of the Product: Product is used in automotive sectors which is CNC machined prototype parts used in design and research field for research and development purpose'


            endUseCell.style.whiteSpace =
              'normal'

            endUseCell.style.overflow =
              'visible'

            endUseCell.style.textOverflow =
              'clip'

            endUseCell.style.height =
              'auto'

            endUseCell.style.minHeight =
              '48px'

            endUseCell.style.lineHeight =
              '1.45'

            endUseCell.style.wordBreak =
              'normal'

            endUseCell.style.overflowWrap =
              'break-word'


            if (endUseRow) {
              endUseRow.style.height =
                'auto'
            }
          }
        }

      }, [
        documentData,
        analysisResult,
        onFieldChange,
        onLineItemChange,
      ])


      return (
        <>
          <style>
            {`
              @keyframes exportflowMissingBlink {
                0%, 100% {
                  box-shadow:
                    inset 0 0 0 2px rgba(255, 40, 40, 0.35),
                    0 0 0 rgba(255, 40, 40, 0);
                }

                50% {
                  box-shadow:
                    inset 0 0 0 3px rgba(255, 30, 30, 1),
                    0 0 12px rgba(255, 30, 30, 0.7);
                }
              }

              .customViewsHost .dynamic-editable {
                cursor: text;
                min-width: 28px;
              }

              .customViewsHost .dynamic-editable:focus {
                outline: 2px solid #3b82f6;
                outline-offset: -2px;
              }

              .customViewsHost .missing-cell {
                animation:
                  exportflowMissingBlink
                  0.9s
                  ease-in-out
                  infinite;
              }

              .customViewsHost .inline-dynamic-field {
                display: inline-block;
                min-width: 70px;
                padding: 0 3px;
              }

              #fixedScometView .scomet-main tr:last-child td {
                white-space: normal !important;
                overflow: visible !important;
                text-overflow: clip !important;
                height: auto !important;
                min-height: 48px !important;
                line-height: 1.45 !important;
                overflow-wrap: break-word !important;
                word-break: normal !important;
              }
            `}
          </style>

          <div
            ref={setHostRef}
            className="customViewsHost"
            dangerouslySetInnerHTML={{
              __html:
                customViewsHtml,
            }}
          />
        </>
      )
    }
  )


export default CustomSheetViews