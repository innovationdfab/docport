export default function HomePage({
  pdfFile,
  progress,
  status,
  onFileChange,
  onAnalyze,
  onGenerate,
}) {
  const openPicker = () => {
    document
      .getElementById('pdf')
      ?.click()
  }

  const handleUploadKeyDown = (
    event
  ) => {
    if (
      event.key === 'Enter' ||
      event.key === ' '
    ) {
      event.preventDefault()
      openPicker()
    }
  }

  return (
    <>
      <style>
        {`
          /* =====================================================
             DOCPORT HOME PAGE
             ===================================================== */

          .docportHome {
            position: relative;
            min-height: 100vh;
            overflow: hidden;
            isolation: isolate;
          }


          /* =====================================================
             AMBIENT BACKGROUND
             ===================================================== */

          .docportHome::before {
            content: "";
            position: fixed;
            width: 460px;
            height: 460px;
            left: -170px;
            top: 15%;
            border-radius: 50%;
            background:
              radial-gradient(
                circle,
                rgba(78, 116, 255, .14),
                transparent 68%
              );
            filter: blur(8px);
            pointer-events: none;
            z-index: -2;
            animation:
              docportOrbOne
              11s
              ease-in-out
              infinite;
          }

          .docportHome::after {
            content: "";
            position: fixed;
            width: 520px;
            height: 520px;
            right: -220px;
            bottom: -190px;
            border-radius: 50%;
            background:
              radial-gradient(
                circle,
                rgba(119, 76, 255, .12),
                transparent 68%
              );
            pointer-events: none;
            z-index: -2;
            animation:
              docportOrbTwo
              13s
              ease-in-out
              infinite;
          }


          @keyframes docportOrbOne {
            0%,
            100% {
              transform:
                translate3d(
                  0,
                  0,
                  0
                );
            }

            50% {
              transform:
                translate3d(
                  45px,
                  -28px,
                  0
                );
            }
          }


          @keyframes docportOrbTwo {
            0%,
            100% {
              transform:
                translate3d(
                  0,
                  0,
                  0
                );
            }

            50% {
              transform:
                translate3d(
                  -45px,
                  -35px,
                  0
                );
            }
          }


          /* =====================================================
             TOP BRAND
             ===================================================== */

          .docportTop {
            animation:
              docportTopEnter
              .7s
              cubic-bezier(
                .22,
                1,
                .36,
                1
              )
              both;
          }


          .docportBrand {
            display: flex;
            align-items: center;
            gap: 15px;
          }


          .docportLogo {
            position: relative;
            display: grid;
            place-items: center;
            overflow: hidden;
            font-size: 15px;
            font-weight: 850;
            letter-spacing: -.5px;
            box-shadow:
              0 16px 38px
              rgba(
                72,
                96,
                255,
                .18
              );
            transition:
              transform .3s ease,
              box-shadow .3s ease;
          }


          .docportLogo::before {
            content: "";
            position: absolute;
            width: 38%;
            height: 160%;
            top: -30%;
            left: -70%;
            transform:
              rotate(18deg);
            background:
              rgba(
                255,
                255,
                255,
                .32
              );
            filter:
              blur(5px);
            animation:
              docportLogoShine
              5.5s
              ease-in-out
              infinite;
          }


          .docportBrand:hover
          .docportLogo {
            transform:
              translateY(-3px)
              rotate(-2deg);
            box-shadow:
              0 20px 46px
              rgba(
                72,
                96,
                255,
                .28
              );
          }


          .docportBrand h2 {
            margin: 0;
            font-size: 23px;
            font-weight: 820;
            letter-spacing: -.5px;
          }


          .docportBrand small {
            display: block;
            margin-top: 4px;
            opacity: .72;
            font-size: 12px;
            letter-spacing: .01em;
          }


          @keyframes docportLogoShine {
            0%,
            60% {
              left: -70%;
            }

            82%,
            100% {
              left: 145%;
            }
          }


          /* =====================================================
             HERO
             ===================================================== */

          .docportHero {
            position: relative;
            animation:
              docportHeroEnter
              .8s
              .08s
              cubic-bezier(
                .22,
                1,
                .36,
                1
              )
              both;
          }


          .docportEyebrow {
            display: inline-flex;
            align-items: center;
            gap: 10px;
            letter-spacing: .19em;
          }


          .docportEyebrow::before,
          .docportEyebrow::after {
            content: "";
            width: 20px;
            height: 1px;
            background:
              currentColor;
            opacity: .5;
          }


          .docportHero h1 {
            letter-spacing: -2.4px;
          }


          .docportHero h1
          .docportAccent {
            position: relative;
            color: #7697ff;
          }


          .docportHero h1
          .docportAccent::after {
            content: "";
            position: absolute;
            left: 4%;
            right: 4%;
            bottom: -8px;
            height: 2px;
            border-radius: 10px;
            background:
              linear-gradient(
                90deg,
                transparent,
                rgba(
                  105,
                  128,
                  255,
                  .8
                ),
                transparent
              );
            transform:
              scaleX(0);
            transform-origin:
              center;
            animation:
              docportUnderline
              .8s
              .8s
              ease
              forwards;
          }


          .docportHero p {
            max-width: 810px;
            margin-left: auto;
            margin-right: auto;
            line-height: 1.7;
          }


          @keyframes docportUnderline {
            to {
              transform:
                scaleX(1);
            }
          }


          /* =====================================================
             WORKFLOW CONTAINER
             ===================================================== */

          .docportFlow {
            position: relative;
            overflow: hidden;
            animation:
              docportFlowEnter
              .85s
              .18s
              cubic-bezier(
                .22,
                1,
                .36,
                1
              )
              both;
          }


          .docportFlow::before {
            content: "";
            position: absolute;
            inset: 0;
            pointer-events: none;
            border-radius: inherit;
            background:
              linear-gradient(
                110deg,
                transparent 0%,
                rgba(
                  255,
                  255,
                  255,
                  .02
                ) 40%,
                rgba(
                  255,
                  255,
                  255,
                  .045
                ) 50%,
                rgba(
                  255,
                  255,
                  255,
                  .02
                ) 60%,
                transparent 100%
              );
            transform:
              translateX(-100%);
            animation:
              docportPanelShine
              8s
              ease-in-out
              infinite;
          }


          @keyframes docportPanelShine {
            0%,
            65% {
              transform:
                translateX(-100%);
            }

            90%,
            100% {
              transform:
                translateX(100%);
            }
          }


          /* =====================================================
             UPLOAD CARD
             ===================================================== */

          .docportUpload {
            position: relative;
            overflow: hidden;
            transition:
              transform .28s ease,
              border-color .28s ease,
              box-shadow .28s ease,
              background .28s ease;
          }


          .docportUpload:hover {
            transform:
              translateY(-5px);
            border-color:
              rgba(
                65,
                132,
                255,
                .92
              );
            box-shadow:
              0 20px 55px
              rgba(
                0,
                0,
                0,
                .16
              );
          }


          .docportUpload::after {
            content: "";
            position: absolute;
            width: 150px;
            height: 150px;
            border-radius: 50%;
            left: -90px;
            bottom: -95px;
            background:
              rgba(
                54,
                126,
                255,
                .08
              );
            transition:
              transform .45s ease;
          }


          .docportUpload:hover::after {
            transform:
              scale(1.5);
          }


          .docportFileIcon {
            transition:
              transform .28s ease,
              box-shadow .28s ease;
          }


          .docportUpload:hover
          .docportFileIcon {
            transform:
              translateY(-4px)
              rotate(-3deg);
            box-shadow:
              0 14px 32px
              rgba(
                0,
                0,
                0,
                .16
              );
          }


          .docportUploadText {
            position: relative;
            z-index: 2;
            min-width: 0;
          }


          .docportUploadText b {
            display: block;
          }


          .docportUploadText span {
            display: block;
          }


          .docportFilename {
            position: relative;
            display: inline-flex;
            align-items: center;
            gap: 7px;
          }


          .docportFilename.selected::before {
            content: "";
            width: 6px;
            height: 6px;
            flex: 0 0 6px;
            border-radius: 50%;
            background:
              #46e7b1;
            box-shadow:
              0 0 10px
              rgba(
                70,
                231,
                177,
                .7
              );
          }


          .docportUploadArrow {
            position: relative;
            z-index: 2;
            margin-left: auto;
            font-size: 22px;
            opacity: 0;
            transform:
              translateX(-8px);
            transition:
              opacity .25s ease,
              transform .25s ease;
          }


          .docportUpload:hover
          .docportUploadArrow {
            opacity: .8;
            transform:
              translateX(0);
          }


          /* =====================================================
             ACTION BUTTONS
             ===================================================== */

          .docportAction {
            position: relative;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 13px;
            transition:
              transform .28s ease,
              border-color .28s ease,
              box-shadow .28s ease,
              background .28s ease;
          }


          .docportAction:hover {
            transform:
              translateY(-5px);
            border-color:
              rgba(
                89,
                126,
                255,
                .55
              );
            box-shadow:
              0 20px 55px
              rgba(
                0,
                0,
                0,
                .16
              );
          }


          .docportAction::before {
            content: "";
            position: absolute;
            top: -30%;
            left: -95px;
            width: 52px;
            height: 165%;
            transform:
              rotate(16deg);
            background:
              linear-gradient(
                90deg,
                transparent,
                rgba(
                  255,
                  255,
                  255,
                  .12
                ),
                transparent
              );
            transition:
              left .65s ease;
          }


          .docportAction:hover::before {
            left:
              calc(
                100% + 60px
              );
          }


          .docportActionIcon {
            position: relative;
            z-index: 2;
            display: grid;
            place-items: center;
            font-size: 20px;
            transition:
              transform .28s ease;
          }


          .docportAction:hover
          .docportActionIcon {
            transform:
              scale(1.14)
              rotate(-5deg);
          }


          .docportActionContent {
            position: relative;
            z-index: 2;
            display: flex;
            flex-direction: column;
            gap: 5px;
            text-align: left;
          }


          .docportActionContent strong {
            font-size: 17px;
          }


          .docportActionContent small {
            font-size: 11px;
            opacity: .55;
            font-weight: 500;
          }


          .docportActionArrow {
            position: relative;
            z-index: 2;
            font-size: 20px;
            opacity: 0;
            transform:
              translateX(-8px);
            transition:
              opacity .25s ease,
              transform .25s ease;
          }


          .docportAction:hover
          .docportActionArrow {
            opacity: .75;
            transform:
              translateX(0);
          }


          /* =====================================================
             PROGRESS AREA
             ===================================================== */

          .docportProgressWrap {
            animation:
              docportStatusEnter
              .4s
              ease
              both;
          }


          .docportProgressWrap
          .progress {
            overflow: hidden;
          }


          .docportProgressWrap
          .progress i {
            position: relative;
            overflow: hidden;
            transition:
              width
              .45s
              cubic-bezier(
                .22,
                1,
                .36,
                1
              );
          }


          .docportProgressWrap
          .progress i::after {
            content: "";
            position: absolute;
            inset: 0;
            background:
              linear-gradient(
                90deg,
                transparent,
                rgba(
                  255,
                  255,
                  255,
                  .36
                ),
                transparent
              );
            transform:
              translateX(-100%);
            animation:
              docportProgressSweep
              1.8s
              linear
              infinite;
          }


          @keyframes docportProgressSweep {
            to {
              transform:
                translateX(100%);
            }
          }


          /* =====================================================
             SMALL FOOTER
             ===================================================== */

          .docportCapabilities {
            display: flex;
            justify-content: center;
            align-items: center;
            flex-wrap: wrap;
            gap: 9px;
            margin-top: 22px;
            font-size: 10.5px;
            letter-spacing: .045em;
            opacity: .38;
            animation:
              docportHeroEnter
              .8s
              .3s
              ease
              both;
          }


          .docportCapabilities i {
            width: 3px;
            height: 3px;
            border-radius: 50%;
            background:
              currentColor;
            opacity: .6;
          }


          /* =====================================================
             ENTRY ANIMATIONS
             ===================================================== */

          @keyframes docportTopEnter {
            from {
              opacity: 0;
              transform:
                translateY(-14px);
            }

            to {
              opacity: 1;
              transform:
                translateY(0);
            }
          }


          @keyframes docportHeroEnter {
            from {
              opacity: 0;
              transform:
                translateY(18px);
            }

            to {
              opacity: 1;
              transform:
                translateY(0);
            }
          }


          @keyframes docportFlowEnter {
            from {
              opacity: 0;
              transform:
                translateY(22px)
                scale(.985);
            }

            to {
              opacity: 1;
              transform:
                translateY(0)
                scale(1);
            }
          }


          @keyframes docportStatusEnter {
            from {
              opacity: 0;
              transform:
                translateY(6px);
            }

            to {
              opacity: 1;
              transform:
                translateY(0);
            }
          }


          /* =====================================================
             MOBILE
             ===================================================== */

          @media (
            max-width: 850px
          ) {
            .docportHero h1 {
              letter-spacing:
                -1.2px;
            }

            .docportAction {
              justify-content:
                flex-start;
              padding-left:
                25px;
            }

            .docportActionArrow,
            .docportUploadArrow {
              margin-left: auto;
              opacity: .65;
              transform: none;
            }

            .docportCapabilities {
              padding:
                0 16px 22px;
            }
          }


          /* =====================================================
             ACCESSIBILITY
             ===================================================== */

          @media (
            prefers-reduced-motion:
            reduce
          ) {
            .docportHome *,
            .docportHome *::before,
            .docportHome *::after {
              animation-duration:
                .001ms !important;

              animation-iteration-count:
                1 !important;

              transition-duration:
                .001ms !important;
            }
          }
        `}
      </style>


      <div
        id="homePage"
        className="page active"
      >
        <main
          className="home docportHome"
        >

          {/* ================================================
              TOP BRAND
              ================================================ */}

          <div
            className="top docportTop"
          >
            <div
              className="brand docportBrand"
            >
              <div
                className="logo docportLogo"
              >
                DP
              </div>

              <div>
                <h2>
                  DocPort
                </h2>

                <small>
                  Invoice-to-Export Documentation
                </small>
              </div>
            </div>

            {/*
              Removed:
              Exact 6-sheet Excel configured
            */}
          </div>


          {/* ================================================
              HERO
              ================================================ */}

          <div
            className="hero docportHero"
          >
            <div
              className="eyebrow docportEyebrow"
            >
              EXPORT DOCUMENT AUTOMATION
            </div>

            <h1>
              Upload.
              <span
                className="docportAccent"
              >
                {' '}Analyze.{' '}
              </span>
              Generate.
            </h1>

            <p>
              Convert commercial invoice PDFs into
              structured export documents, review
              extracted shipment data and generate
              a ready-to-use Excel workbook.
            </p>
          </div>


          {/* ================================================
              WORKFLOW
              ================================================ */}

          <section
            className="flow docportFlow"
            aria-label="Export document workflow"
          >
            <div className="steps">


              {/* ============================================
                  UPLOAD
                  ============================================ */}

              <div
                className="upload docportUpload"
                role="button"
                tabIndex={0}
                onClick={
                  openPicker
                }
                onKeyDown={
                  handleUploadKeyDown
                }
              >
                <div
                  className="fileIcon docportFileIcon"
                  aria-hidden="true"
                >
                  <svg
                    width="29"
                    height="29"
                    viewBox="0 0 24 24"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="1.7"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <path
                      d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"
                    />

                    <path
                      d="M14 2v6h6"
                    />

                    <path
                      d="M8 13h8"
                    />

                    <path
                      d="M8 17h6"
                    />
                  </svg>
                </div>


                <div
                  className="docportUploadText"
                >
                  <b>
                    Upload Commercial Invoice PDF
                  </b>

                  <span>
                    Select the current shipment invoice.
                  </span>


                  <div
                    id="pdfName"
                    className={
                      `filename docportFilename${
                        pdfFile
                          ? ' selected'
                          : ''
                      }`
                    }
                  >
                    {
                      pdfFile
                        ? pdfFile.name
                        : 'No PDF selected'
                    }
                  </div>
                </div>


                <span
                  className="docportUploadArrow"
                  aria-hidden="true"
                >
                  →
                </span>
              </div>


              <input
                id="pdf"
                type="file"
                accept=".pdf,application/pdf"
                hidden
                onChange={(
                  event
                ) =>
                  onFileChange(
                    event
                      .target
                      .files?.[0] ||
                      null
                  )
                }
              />


              {/* ============================================
                  ANALYZE
                  ============================================ */}

              <button
                id="analyzeBtn"
                className="action docportAction"
                type="button"
                onClick={
                  onAnalyze
                }
              >
                <span
                  className="docportActionIcon"
                  aria-hidden="true"
                >
                  ✦
                </span>


                <span
                  className="docportActionContent"
                >
                  <strong>
                    Analyze PDF
                  </strong>

                  <small>
                    Extract shipment data
                  </small>
                </span>


                <span
                  className="docportActionArrow"
                  aria-hidden="true"
                >
                  →
                </span>
              </button>


              {/* ============================================
                  GENERATE
                  ============================================ */}

              <button
                id="generateBtn"
                className="action docportAction"
                type="button"
                onClick={
                  onGenerate
                }
              >
                <span
                  className="docportActionIcon"
                  aria-hidden="true"
                >
                  ▣
                </span>


                <span
                  className="docportActionContent"
                >
                  <strong>
                    Generate Excel
                  </strong>

                  <small>
                    Open the 6-sheet workbook
                  </small>
                </span>


                <span
                  className="docportActionArrow"
                  aria-hidden="true"
                >
                  →
                </span>
              </button>

            </div>


            {/* ================================================
                PROGRESS
                ================================================ */}

            <div
              id="progressWrap"
              className="progressWrap docportProgressWrap"
              style={{
                display:
                  progress > 0
                    ? 'block'
                    : 'none',
              }}
            >
              <div
                className="progress"
                aria-label="PDF analysis progress"
              >
                <i
                  id="bar"
                  style={{
                    width:
                      `${progress}%`,
                  }}
                />
              </div>


              <div
                id="status"
                className="status"
              >
                {status}
              </div>
            </div>
          </section>


          {/* ================================================
              SMALL CAPABILITY LINE
              ================================================ */}

          <div
            className="docportCapabilities"
          >
            <span>
              PDF Extraction
            </span>

            <i />

            <span>
              Data Review
            </span>

            <i />

            <span>
              Manual Editing
            </span>

            <i />

            <span>
              Excel Export
            </span>
          </div>

        </main>
      </div>
    </>
  )
}