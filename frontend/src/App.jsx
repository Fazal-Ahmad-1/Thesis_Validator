import { useMemo, useState, useRef } from 'react'

const UNIVERSITIES = [
  {
    name: 'Arizona State University',
    programs: [
      { profile: 'asu_graduate', label: 'Graduate College — Thesis/Dissertation' },
    ],
  },
  {
    name: 'Amity University Mumbai',
    programs: [
      { profile: 'amity_selected', label: 'B.Sc. — Project Report' },
    ],
  },
  {
    name: 'University of Mumbai',
    programs: [
      { profile: 'mumbai_engineering', label: 'Engineering — Thesis/Dissertation/Project Report' },
    ],
  },
  {
    name: 'University of Kota',
    programs: [
      { profile: 'uok_phd', label: 'Ph.D. — Thesis' },
    ],
  },
  {
    name: 'Guru Ghasidas Vishwavidyalaya',
    programs: [
      { profile: 'ggu_political_science', label: 'MA/PhD Political Science — Dissertation/Thesis' },
    ],
  },
]

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://127.0.0.1:8000'
const API_URL = `${API_BASE_URL}/api/validate`

// ---- Presentation-only helpers -------------------------------------------
// These translate the backend's raw fields into plain language. The
// backend response shape is never modified.

const RULE_TITLES = {
  'Page Size': 'Page size is incorrect',
  'Top Margin': 'Top margin is incorrect',
  'Bottom Margin': 'Bottom margin is incorrect',
  'Left Margin': 'Left margin is incorrect',
  'Right Margin': 'Right margin is incorrect',
  'Font': "Font doesn't match the required style",
  'Line Spacing': 'Line spacing is incorrect',
  'Page Number Field': 'Page numbers are missing',
  'Page Number Alignment': 'Page number position is incorrect',
  'Page Numbering Format': 'Page numbering format could not be verified',
  'Main Text Page Numbering Format': 'Main text page numbering format is incorrect',
  'Main Text Page Numbering Start': 'Main text page numbering does not start correctly',
  'Preliminary Page Numbering Format': 'Preliminary page numbering could not be verified',
}

// Fallback only — used if the backend ever omits "message" for an issue.
function humanizeTitle(rule) {
  return RULE_TITLES[rule] || rule || 'Formatting issue'
}

function humanizeValue(value) {
  if (typeof value !== 'string') return value
  // "8.50 x 14.00 inches" -> "8.50 × 14.00 inches"
  return value.replace(/(\d)\s*x\s*(\d)/gi, '$1 \u00d7 $2')
}

// Only used for the progress-bar/score tint; the score/label text itself
// always comes directly from the backend response.
function scoreColor(score) {
  if (score >= 90) return '#2F6846'
  if (score >= 75) return '#3E6B8A'
  if (score >= 50) return '#8A6D1F'
  return '#A8332B'
}

// Short, student-friendly explanation derived purely from the score
// thresholds already defined by the backend's scoring rule. This is
// presentation text only — it does not alter the score or its label.
function scoreExplanation(score) {
  if (score >= 90) return 'Your document follows the selected formatting requirements well.'
  if (score >= 75) return 'Your document is mostly formatted correctly, but some issues need attention.'
  if (score >= 50) return 'Your document needs several formatting changes before submission.'
  return 'Your document has major formatting issues that should be fixed before submission.'
}

// ---- "Passed Checks" helpers ----------------------------------------------
// Page/Font/Paragraph/Pagination are always evaluated by the backend for
// every current profile, so their absence from blocking_errors reliably
// means "passed". Heading validation is optional in the backend (only
// runs when a profile defines heading rules), so it's only shown as
// passed when it actually appears somewhere in this result's error list
// (blocking or informational) — i.e. there's real evidence it ran.
const CATEGORY_LABELS = {
  Page: 'Page formatting',
  Font: 'Font formatting',
  Paragraph: 'Paragraph formatting',
  Pagination: 'Pagination',
  Heading: 'Heading formatting',
}
const CATEGORY_DISPLAY_ORDER = ['Page', 'Font', 'Paragraph', 'Pagination', 'Heading']


function issueKey(issue) {
  return [
    issue.category || '',
    issue.rule || '',
    issue.message || '',
    issue.expected || '',
    issue.actual || '',
  ].join('|')
}

function getCategoryLabel(category) {
  return CATEGORY_LABELS[category] || category || 'Other'
}

function getIssueFix(issue) {
  if (issue.suggestion) return issue.suggestion

  if (issue.category === 'Font') {
    return 'Select the affected text in Word and apply the required font family and size.'
  }

  if (issue.category === 'Paragraph') {
    return 'Select the affected paragraphs in Word and apply the required line spacing.'
  }

  if (issue.category === 'Page') {
    return 'Open Layout → Page Setup in Word and apply the required page size or margins.'
  }

  if (issue.category === 'Pagination') {
    return 'Check the header/footer and page-number settings in Word against the selected profile.'
  }

  if (issue.category === 'Heading') {
    return 'Select the affected heading and apply the required heading formatting.'
  }

  return null
}

function groupIssues(issues) {
  const groups = new Map()

  issues.forEach((issue) => {
    const key = issueKey(issue)

    if (!groups.has(key)) {
      groups.set(key, {
        ...issue,
        groupedLocations: issue.location ? [issue.location] : [],
        groupedOccurrences:
          typeof issue.occurrences === 'number' ? issue.occurrences : 1,
      })
      return
    }

    const group = groups.get(key)

    if (issue.location && !group.groupedLocations.includes(issue.location)) {
      group.groupedLocations.push(issue.location)
    }

    if (typeof issue.occurrences === 'number') {
      group.groupedOccurrences += issue.occurrences
    } else {
      group.groupedOccurrences += 1
    }
  })

  return Array.from(groups.values()).map((group) => ({
    ...group,
    occurrences: group.groupedOccurrences,
    location:
      group.groupedLocations.length > 0
        ? group.groupedLocations.join(', ')
        : group.location,
  }))
}

function IssueCard({ issue }) {
  const isInfo = issue.severity === 'info'
  const barColor = isInfo ? 'bg-[#8A6D1F]' : 'bg-[#A8332B]'
  const hasExpected = issue.expected !== undefined && issue.expected !== null
  const hasActual = issue.actual !== undefined && issue.actual !== null
  const hasOccurrences = typeof issue.occurrences !== 'undefined' && issue.occurrences > 1

  return (
    <li className="flex gap-3 py-4 first:pt-0 last:pb-0">
      <div className={`mt-1 w-1 shrink-0 rounded-full ${barColor}`} aria-hidden="true" />
      <div className="min-w-0 flex-1">
        {(issue.category || issue.type) && (
          <div className="mb-1.5 flex flex-wrap items-center gap-2">
            {issue.category && (
              <span className="inline-block border border-[#D8D4CB] px-2 py-0.5 text-xs text-[#5B5A54]">
                {issue.category}
              </span>
            )}
            {isInfo && (
              <span className="inline-block border border-[#E4D9B0] bg-[#FBF6E8] px-2 py-0.5 text-xs text-[#8A6D1F]">
                Informational — not a failure
              </span>
            )}
          </div>
        )}

        {/* Message — the prominent "what is wrong" line */}
        <p className="text-[15px] font-medium text-[#1E2A38]">
          {issue.message || humanizeTitle(issue.rule)}
        </p>

        {/* Expected / Actual — clearly labeled, side by side */}
        {(hasExpected || hasActual) && (
          <dl className="mt-2 grid grid-cols-1 gap-x-4 gap-y-1 text-sm sm:grid-cols-2">
            {hasExpected && (
              <div>
                <dt className="text-xs uppercase tracking-wide text-[#8A887E]">Expected</dt>
                <dd className="text-[#1E2A38]">{humanizeValue(issue.expected)}</dd>
              </div>
            )}
            {hasActual && (
              <div>
                <dt className="text-xs uppercase tracking-wide text-[#8A887E]">Actual</dt>
                <dd className="text-[#1E2A38]">{humanizeValue(issue.actual)}</dd>
              </div>
            )}
          </dl>
        )}

        {/* Location / occurrences — secondary metadata, de-emphasized */}
        {(issue.location || hasOccurrences) && (
          <dl className="mt-1.5 space-y-0.5 text-xs text-[#8A887E]">
            {issue.location && (
              <div className="flex gap-1.5">
                <dt>Found in:</dt>
                <dd>{issue.location}</dd>
              </div>
            )}
            {hasOccurrences && (
              <div className="flex gap-1.5">
                <dt>Occurs:</dt>
                <dd>{issue.occurrences} times</dd>
              </div>
            )}
          </dl>
        )}

        {/* Suggestion — visually separated as the actionable recommendation */}
        {getIssueFix(issue) && (
          <p className="mt-2.5 border-l-2 border-[#1F3A5F] bg-[#EEF2F6] px-3 py-2 text-sm text-[#1F3A5F]">
            <span className="font-medium">How to fix: </span>
            {getIssueFix(issue)}
          </p>
        )}
      </div>
    </li>
  )
}

function App() {
  const [university, setUniversity] = useState('')
  const [profileName, setProfileName] = useState('')
  const [selectedFile, setSelectedFile] = useState(null)
  const [fileError, setFileError] = useState('')
  const [isDragging, setIsDragging] = useState(false)
  const [isLoading, setIsLoading] = useState(false)
  const [result, setResult] = useState(null)
  const [requestError, setRequestError] = useState('')
  const [issueFilter, setIssueFilter] = useState('all')
  const [categoryFilter, setCategoryFilter] = useState('all')
  const [showValidator, setShowValidator] = useState(false)
  const [showPrivacyPolicy, setShowPrivacyPolicy] = useState(false)
  const fileInputRef = useRef(null)

  const selectedUniversity = UNIVERSITIES.find((u) => u.name === university)

  const handleFile = (file) => {
    setResult(null)
    setRequestError('')

    if (!file) return

    if (!file.name.toLowerCase().endsWith('.docx')) {
      setSelectedFile(null)
      setFileError('Only .docx files are supported. Please choose a Word document.')
      return
    }

    setFileError('')
    setSelectedFile(file)
  }

  const handleInputChange = (event) => {
    handleFile(event.target.files?.[0] || null)
  }

  const handleDrop = (event) => {
    event.preventDefault()
    setIsDragging(false)
    handleFile(event.dataTransfer.files?.[0] || null)
  }

  const handleRemoveFile = () => {
    setSelectedFile(null)
    setFileError('')
    if (fileInputRef.current) fileInputRef.current.value = ''
  }

  const handleUniversityChange = (event) => {
    setUniversity(event.target.value)
    // A university change invalidates any previously selected program,
    // since program options depend on the chosen university.
    setProfileName('')
    setResult(null)
    setRequestError('')
  }

  const handleProfileChange = (event) => {
    setProfileName(event.target.value)
    setResult(null)
    setRequestError('')
  }

  const handleValidate = async () => {
    if (!selectedFile || !profileName) return

    setIsLoading(true)
    setRequestError('')
    setResult(null)
    setIssueFilter('all')
    setCategoryFilter('all')

    try {
      const formData = new FormData()
      formData.append('file', selectedFile)
      formData.append('profile_name', profileName)

      const response = await fetch(API_URL, {
        method: 'POST',
        body: formData,
      })

      if (!response.ok) {
        let message = `Validation failed (HTTP ${response.status}).`
        try {
          const errorBody = await response.json()
          if (errorBody?.detail) message = errorBody.detail
        } catch {
          // Response wasn't JSON; keep the default message.
        }
        setRequestError(message)
        return
      }

      setResult(await response.json())
    } catch {
      setRequestError(
        'Could not reach the validation server. Please make sure the backend is running and try again.'
      )
    } finally {
      setIsLoading(false)
    }
  }

  const handleReset = () => {
    setResult(null)
    setSelectedFile(null)
    setFileError('')
    setRequestError('')
    setIssueFilter('all')
    setCategoryFilter('all')
    setUniversity('')
    setProfileName('')
    if (fileInputRef.current) fileInputRef.current.value = ''
  }

  const canValidate = Boolean(selectedFile) && Boolean(profileName) && !isLoading

  const blockingErrors = result?.blocking_errors ?? []
  const informationalNotices = result?.informational_notices ?? []
  const blockingCount = result?.blocking_error_count ?? blockingErrors.length
  const infoCount = result?.informational_notice_count ?? informationalNotices.length
  const allResultIssues = result?.errors ?? []

  // Use the backend's own score fields directly — no recalculation here.
  const score =
    typeof result?.score === 'number'
      ? result.score
      : typeof result?.validation_score === 'number'
        ? result.validation_score
        : null
  const scoreLabel = result?.score_label ?? result?.validation_score_label ?? null

  // The backend explicitly tells us whether each check is passed,
  // needs attention, or is not enforced by the selected profile.
  // Keep a small fallback for older API responses that do not contain
  // the new `checks` metadata.
  const checkSummary = useMemo(() => {
    if (result?.checks && typeof result.checks === 'object') {
      return CATEGORY_DISPLAY_ORDER
        .filter((category) => result.checks[category])
        .map((category) => ({
          category,
          status: result.checks[category].status,
          reason: result.checks[category].reason || null,
        }))
    }

    const categoriesWithBlockingIssues = new Set(
      blockingErrors.map((e) => e.category).filter(Boolean)
    )
    const categoriesObservedThisRun = new Set(
      allResultIssues.map((e) => e.category).filter(Boolean)
    )
    const fallbackCategories = new Set([
      'Page',
      'Font',
      'Paragraph',
      'Pagination',
      ...categoriesObservedThisRun,
    ])

    return CATEGORY_DISPLAY_ORDER
      .filter((category) => fallbackCategories.has(category))
      .map((category) => ({
        category,
        status: categoriesWithBlockingIssues.has(category)
          ? 'needs_attention'
          : 'passed',
        reason: null,
      }))
  }, [result, blockingErrors, allResultIssues])

  const passedCheckCount = checkSummary.filter(
    (check) => check.status === 'passed'
  ).length

  const issueCategories = useMemo(() => {
    const categories = new Set(
      allResultIssues.map((issue) => issue.category).filter(Boolean)
    )

    return CATEGORY_DISPLAY_ORDER.filter((category) => categories.has(category))
  }, [allResultIssues])

  const filteredBlockingErrors = useMemo(() => {
    if (issueFilter === 'info') return []

    let issues = blockingErrors

    if (categoryFilter !== 'all') {
      issues = issues.filter((issue) => issue.category === categoryFilter)
    }

    return issues
  }, [blockingErrors, issueFilter, categoryFilter])

  const filteredInformationalNotices = useMemo(() => {
    if (issueFilter === 'errors') return []

    let issues = informationalNotices

    if (categoryFilter !== 'all') {
      issues = issues.filter((issue) => issue.category === categoryFilter)
    }

    return issues
  }, [informationalNotices, issueFilter, categoryFilter])

  const blockingByCategory = CATEGORY_DISPLAY_ORDER.map((category) => ({
    category,
    issues: groupIssues(
      filteredBlockingErrors.filter((issue) => issue.category === category)
    ),
  })).filter((group) => group.issues.length > 0)

  const groupedInfoIssues = groupIssues(filteredInformationalNotices)

  const visibleBlockingCount = filteredBlockingErrors.length
  const visibleInfoCount = filteredInformationalNotices.length
  const hasVisibleIssues = visibleBlockingCount > 0 || visibleInfoCount > 0

  // Only render profile-context fields the API actually returned —
  // never show an empty/undefined field.
  const profileFields = [
    { label: 'University', value: result?.university },
    { label: 'Program', value: result?.program },
    { label: 'Document Type', value: result?.document_type },
  ].filter((field) => field.value)

  const openValidator = () => {
    setShowValidator(true)
    window.setTimeout(() => {
      document.getElementById('validator')?.scrollIntoView({ behavior: 'smooth', block: 'start' })
    }, 0)
  }

  const goHome = () => {
    setShowValidator(false)
    setShowPrivacyPolicy(false)
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  const openPrivacyPolicy = () => {
    setShowValidator(false)
    setShowPrivacyPolicy(true)
    window.scrollTo({ top: 0, behavior: 'smooth' })
  }

  if (showPrivacyPolicy) {
    return (
      <div className="min-h-screen bg-[#F6F5F1] text-[#1E2A38]">
        <header className="border-b border-[#D8D4CB] bg-white/90">
          <div className="mx-auto flex max-w-6xl items-center justify-between px-5 py-4 sm:px-8">
            <button onClick={goHome} className="font-serif text-xl font-semibold tracking-tight">
              Thesis Validator
            </button>
            <button
              onClick={goHome}
              className="text-sm text-[#1F3A5F] hover:underline"
            >
              ← Back to home
            </button>
          </div>
        </header>

        <main>
          <section className="border-b border-[#D8D4CB] bg-white">
            <div className="mx-auto max-w-4xl px-5 py-16 sm:px-8 sm:py-20">
              <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">
                Legal
              </p>
              <h1 className="mt-3 font-serif text-4xl leading-tight sm:text-5xl">
                Privacy Policy
              </h1>
              <p className="mt-4 text-sm text-[#8A887E]">
                Last updated: September 17, 2026
              </p>
              <p className="mt-6 max-w-3xl text-base leading-7 text-[#5B5A54]">
                This Privacy Policy explains how Thesis Validator handles information when you
                use the current web application to check the formatting of an academic DOCX file.
                It is written for the current MVP and will be updated if the product's data
                practices change.
              </p>
            </div>
          </section>

          <section>
            <div className="mx-auto max-w-4xl px-5 py-12 sm:px-8 sm:py-16">
              <div className="space-y-10">
                {[
                  [
                    '1. Information you provide',
                    [
                      'The current validator lets you select a university/program and upload a DOCX file for formatting validation.',
                      'The current MVP does not require an account, so it does not intentionally collect a name, email address, password, or profile information through an account system.',
                    ],
                  ],
                  [
                    '2. How uploaded documents are processed',
                    [
                      'An uploaded DOCX is sent to the validation backend so the application can inspect its document structure and compare supported formatting properties with the selected profile.',
                      'The application is designed to analyze formatting rather than evaluate the academic quality, originality, grammar, or research content of your work.',
                    ],
                  ],
                  [
                    '3. Document storage and deletion',
                    [
                      'The current MVP does not intentionally store uploaded documents as part of the validation process.',
                      'During processing, the backend writes the upload to a temporary server-side file. That temporary file is removed after processing, including when validation encounters an error.',
                      'This policy describes the application behavior implemented for the current MVP. Hosting infrastructure, network providers, or other services may have their own technical logs or retention practices outside the application itself.',
                    ],
                  ],
                  [
                    '4. Validation results',
                    [
                      'The validator returns formatting results to your browser, including detected issues, expected values, actual values, and informational notices where applicable.',
                      'The current MVP does not provide an account-based history feature for saving your validation results.',
                    ],
                  ],
                  [
                    '5. Information collected automatically',
                    [
                      'The current MVP does not intentionally add advertising trackers, analytics software, or third-party profiling tools to the application code.',
                      'Your web browser, hosting provider, or network infrastructure may still generate technical logs such as IP addresses, request times, browser information, or error records according to their own policies.',
                    ],
                  ],
                  [
                    '6. Third-party services',
                    [
                      'The application may be hosted using third-party infrastructure. Those providers operate their own systems and privacy policies.',
                      'If additional services such as analytics, authentication, payments, or external document-processing APIs are introduced, this policy will be updated to describe the relevant data handling before those features are used.',
                    ],
                  ],
                  [
                    '7. Security',
                    [
                      'The validator uses server-side temporary processing for uploaded documents and removes temporary files after processing.',
                      'No online service can guarantee absolute security. You should avoid uploading documents containing information that you are not comfortable submitting to an online service.',
                    ],
                  ],
                  [
                    '8. Your choices',
                    [
                      'Because the current MVP does not require an account, there is no account profile to manage or delete.',
                      'You can choose not to upload a document. You should also avoid submitting confidential material unless you are comfortable with the service receiving it for processing.',
                    ],
                  ],
                  [
                    '9. Changes to this policy',
                    [
                      'This policy may be updated when the application\'s functionality, storage practices, hosting arrangements, or data handling changes. The “Last updated” date at the top of this page will be changed when material updates are made.',
                    ],
                  ],
                  [
                    '10. Contact',
                    [
                      'For privacy questions or concerns about the current application, please use the contact method provided by the product operator. A dedicated support/contact address can be added here before public launch.',
                    ],
                  ],
                ].map(([title, paragraphs]) => (
                  <section key={title}>
                    <h2 className="font-serif text-2xl">{title}</h2>
                    <div className="mt-3 space-y-3 text-sm leading-7 text-[#5B5A54]">
                      {paragraphs.map((paragraph) => (
                        <p key={paragraph}>{paragraph}</p>
                      ))}
                    </div>
                  </section>
                ))}

                <div className="border border-[#D8D4CB] bg-white p-5 sm:p-6">
                  <p className="text-sm font-semibold text-[#1E2A38]">
                    Important note about the current MVP
                  </p>
                  <p className="mt-2 text-sm leading-6 text-[#5B5A54]">
                    This policy is intentionally limited to features that currently exist. It is
                    not a promise that future authentication, payments, analytics, storage,
                    integrations, or other features will operate the same way. The policy should
                    be reviewed and updated before public launch and whenever data practices change.
                  </p>
                </div>
              </div>
            </div>
          </section>
        </main>

        <footer className="border-t border-[#D8D4CB] bg-white">
          <div className="mx-auto flex max-w-6xl flex-col gap-3 px-5 py-6 text-sm text-[#8A887E] sm:flex-row sm:items-center sm:justify-between sm:px-8">
            <span>© 2026 Thesis Validator</span>
            <button onClick={goHome} className="text-left hover:text-[#1F3A5F] hover:underline sm:text-right">
              Back to home
            </button>
          </div>
        </footer>
      </div>
    )
  }

  if (!showValidator) {
    return (
      <div className="min-h-screen bg-[#F6F5F1] text-[#1E2A38]">
        <header className="border-b border-[#D8D4CB] bg-white/90">
          <div className="mx-auto flex max-w-6xl items-center justify-between px-5 py-4 sm:px-8">
            <button onClick={goHome} className="font-serif text-xl font-semibold tracking-tight">
              Thesis Validator
            </button>
            <button
              onClick={openValidator}
              className="bg-[#1F3A5F] px-4 py-2 text-sm font-medium text-white transition hover:bg-[#172D49]"
            >
              Check My Document
            </button>
          </div>
        </header>

        <main>
          {/* Hero */}
          <section className="border-b border-[#D8D4CB]">
            <div className="mx-auto grid max-w-6xl gap-10 px-5 py-16 sm:px-8 sm:py-24 lg:grid-cols-[1.15fr_0.85fr] lg:items-center">
              <div>
                <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">
                  University-specific formatting checks
                </p>
                <h1 className="mt-4 max-w-3xl font-serif text-4xl leading-tight sm:text-6xl">
                  Check your thesis formatting before you submit.
                </h1>
                <p className="mt-6 max-w-2xl text-base leading-7 text-[#5B5A54] sm:text-lg">
                  Upload your Word document and review its formatting against the requirements configured for your university and program.
                </p>
                <div className="mt-8 flex flex-col gap-3 sm:flex-row sm:items-center">
                  <button
                    onClick={openValidator}
                    className="bg-[#1F3A5F] px-6 py-3.5 text-sm font-semibold text-white transition hover:bg-[#172D49]"
                  >
                    Check My Document →
                  </button>
                  <span className="text-sm text-[#8A887E]">DOCX files · Formatting report · No account required in this MVP</span>
                </div>
              </div>

              <div className="border border-[#D8D4CB] bg-white p-5 shadow-sm sm:p-6">
                <div className="flex items-center justify-between border-b border-[#EDEBE4] pb-4">
                  <div>
                    <p className="text-xs uppercase tracking-wide text-[#8A887E]">Example report</p>
                    <p className="mt-1 font-medium">Formatting review</p>
                  </div>
                  <span className="border border-[#CFE3D6] bg-[#F1F7F3] px-2 py-1 text-xs font-medium text-[#2F6846]">82 / 100</span>
                </div>
                <div className="mt-5 space-y-3">
                  {[['✓', 'Page formatting', 'Passed'], ['✕', 'Font formatting', 'Needs attention'], ['✓', 'Paragraph formatting', 'Passed'], ['✕', 'Pagination', 'Needs attention']].map(([icon, label, status]) => (
                    <div key={label} className="flex items-center justify-between border border-[#EDEBE4] px-3 py-3">
                      <span className="flex items-center gap-2 text-sm"><span className="font-semibold">{icon}</span>{label}</span>
                      <span className="text-xs text-[#8A887E]">{status}</span>
                    </div>
                  ))}
                </div>
                <p className="mt-5 border-l-2 border-[#1F3A5F] bg-[#EEF2F6] px-3 py-2 text-sm text-[#1F3A5F]">
                  See what needs changing, why it matters, and what to fix in Word.
                </p>
              </div>
            </div>
          </section>

          {/* Problem */}
          <section className="border-b border-[#D8D4CB] bg-white">
            <div className="mx-auto max-w-6xl px-5 py-16 sm:px-8">
              <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">Before submission</p>
              <h2 className="mt-3 max-w-2xl font-serif text-3xl sm:text-4xl">
                Small formatting mistakes can be easy to miss.
              </h2>
              <div className="mt-8 grid gap-4 sm:grid-cols-2 lg:grid-cols-5">
                {['Incorrect margins', 'Wrong font or size', 'Incorrect line spacing', 'Missing page numbers', 'Heading formatting issues'].map((item) => (
                  <div key={item} className="border border-[#D8D4CB] p-5">
                    <span className="text-[#A8332B]">✕</span>
                    <p className="mt-3 text-sm font-medium">{item}</p>
                  </div>
                ))}
              </div>
            </div>
          </section>

          {/* How it works */}
          <section className="border-b border-[#D8D4CB]">
            <div className="mx-auto max-w-6xl px-5 py-16 sm:px-8">
              <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">How it works</p>
              <h2 className="mt-3 font-serif text-3xl sm:text-4xl">Three simple steps.</h2>
              <div className="mt-10 grid gap-8 md:grid-cols-3">
                {[
                  ['01', 'Select your university', 'Choose the university and program whose formatting requirements you want to check.'],
                  ['02', 'Upload your DOCX', 'Upload your thesis, dissertation, or project report as a Word document.'],
                  ['03', 'Review your report', 'See your score, detected issues, expected values, and practical fix suggestions.'],
                ].map(([number, title, text]) => (
                  <div key={number}>
                    <p className="text-sm font-semibold text-[#1F3A5F]">{number}</p>
                    <h3 className="mt-2 text-lg font-semibold">{title}</h3>
                    <p className="mt-2 text-sm leading-6 text-[#5B5A54]">{text}</p>
                  </div>
                ))}
              </div>
            </div>
          </section>

          {/* Coverage */}
          <section className="border-b border-[#D8D4CB] bg-white">
            <div className="mx-auto max-w-6xl px-5 py-16 sm:px-8">
              <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">Current coverage</p>
              <h2 className="mt-3 font-serif text-3xl sm:text-4xl">Supported university profiles</h2>
              <div className="mt-8 overflow-hidden border border-[#D8D4CB]">
                <div className="hidden grid-cols-[1.2fr_1fr_1.4fr] border-b border-[#D8D4CB] bg-[#FAFAF8] px-5 py-3 text-xs font-semibold uppercase tracking-wide text-[#8A887E] sm:grid">
                  <span>University</span><span>Program</span><span>Document</span>
                </div>
                {UNIVERSITIES.map((u) => u.programs.map((program) => (
                  <div key={program.profile} className="grid gap-1 border-b border-[#EDEBE4] px-5 py-4 last:border-b-0 sm:grid-cols-[1.2fr_1fr_1.4fr] sm:gap-0">
                    <span className="font-medium">{u.name}</span>
                    <span className="text-sm text-[#5B5A54]">{program.label.split(' — ')[0]}</span>
                    <span className="text-sm text-[#5B5A54]">{program.label.split(' — ')[1]}</span>
                  </div>
                )))}
              </div>
              <p className="mt-4 text-sm text-[#8A887E]">More university-specific profiles can be added as requirements are verified.</p>
            </div>
          </section>

          {/* What we check / don't check */}
          <section className="border-b border-[#D8D4CB]">
            <div className="mx-auto grid max-w-6xl gap-8 px-5 py-16 sm:px-8 lg:grid-cols-2">
              <div className="border border-[#D8D4CB] bg-white p-6 sm:p-8">
                <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">Currently checked</p>
                <ul className="mt-5 space-y-3 text-sm">
                  {['Page size', 'Margins', 'Font family and size', 'Line spacing', 'Heading formatting where configured', 'Page-number field structure'].map((item) => (
                    <li key={item} className="flex gap-2"><span className="font-semibold text-[#2F6846]">✓</span>{item}</li>
                  ))}
                </ul>
              </div>
              <div className="border border-[#D8D4CB] bg-white p-6 sm:p-8">
                <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">Not currently checked</p>
                <ul className="mt-5 space-y-3 text-sm text-[#5B5A54]">
                  {['Citation correctness', 'Grammar or writing quality', 'Plagiarism', 'Research quality', 'Exact rendered or printed pagination'].map((item) => (
                    <li key={item} className="flex gap-2"><span className="text-[#8A887E]">—</span>{item}</li>
                  ))}
                </ul>
              </div>
            </div>
          </section>

          {/* Privacy & document handling */}
          <section className="border-b border-[#D8D4CB] bg-white">
            <div className="mx-auto max-w-5xl px-5 py-16 sm:px-8">
              <div className="max-w-3xl">
                <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">
                  Privacy &amp; document handling
                </p>
                <h2 className="mt-3 font-serif text-3xl sm:text-4xl">
                  Your research belongs to you.
                </h2>
                <p className="mt-5 text-sm leading-7 text-[#5B5A54]">
                  Academic documents may contain unpublished research and personal information.
                  The validator is designed to analyze document formatting, not evaluate the
                  academic content of your work.
                </p>
              </div>

              <div className="mt-10 grid gap-4 md:grid-cols-3">
                {[
                  [
                    '01',
                    'Temporary processing',
                    'Your DOCX is processed to check its formatting against the university and program profile you select.',
                  ],
                  [
                    '02',
                    'No intentional document storage',
                    'The current MVP does not intentionally store uploaded documents as part of the validation process.',
                  ],
                  [
                    '03',
                    'Temporary files are removed',
                    'The backend uses a temporary processing file and removes it after validation, including when validation encounters an error.',
                  ],
                ].map(([number, title, text]) => (
                  <div key={number} className="border border-[#D8D4CB] bg-[#FAFAF8] p-5">
                    <p className="text-sm font-semibold text-[#1F3A5F]">{number}</p>
                    <h3 className="mt-2 text-base font-semibold">{title}</h3>
                    <p className="mt-2 text-sm leading-6 text-[#5B5A54]">{text}</p>
                  </div>
                ))}
              </div>

              <div className="mt-6 border border-[#D8D4CB] p-5 sm:p-6">
                <p className="text-sm font-semibold text-[#1E2A38]">A few important limits</p>
                <ul className="mt-4 space-y-2 text-sm leading-6 text-[#5B5A54]">
                  <li className="flex gap-2">
                    <span className="text-[#8A887E]">—</span>
                    <span>Do not upload documents containing information you are not comfortable submitting to an online service.</span>
                  </li>
                  <li className="flex gap-2">
                    <span className="text-[#8A887E]">—</span>
                    <span>The validator does not modify your original DOCX file.</span>
                  </li>
                  <li className="flex gap-2">
                    <span className="text-[#8A887E]">—</span>
                    <span>Formatting results are an aid and do not guarantee university approval or submission acceptance.</span>
                  </li>
                </ul>
              </div>
            </div>
          </section>

          {/* FAQ */}
          <section className="border-b border-[#D8D4CB]">
            <div className="mx-auto max-w-4xl px-5 py-16 sm:px-8">
              <p className="text-sm font-medium uppercase tracking-[0.18em] text-[#6B6A63]">FAQ</p>
              <div className="mt-6 divide-y divide-[#D8D4CB] border-y border-[#D8D4CB]">
                {[
                  ['What file types are supported?', 'DOCX Word documents are currently supported.'],
                  ['Does it check my university requirements?', 'It checks against the specific university and program profile selected before upload.'],
                  ['Does it check grammar or plagiarism?', 'No. The current product focuses on document formatting.'],
                  ['Can it guarantee my thesis will be accepted?', 'No. It is a formatting aid and does not replace official university guidelines.'],
                ].map(([question, answer]) => (
                  <details key={question} className="group px-1 py-5">
                    <summary className="cursor-pointer list-none font-medium">{question}<span className="float-right text-[#8A887E]">+</span></summary>
                    <p className="mt-3 max-w-3xl text-sm leading-6 text-[#5B5A54]">{answer}</p>
                  </details>
                ))}
              </div>
            </div>
          </section>

          {/* Final CTA */}
          <section>
            <div className="mx-auto max-w-4xl px-5 py-20 text-center sm:px-8 sm:py-24">
              <h2 className="font-serif text-4xl sm:text-5xl">Ready to check your document?</h2>
              <p className="mx-auto mt-4 max-w-xl text-sm leading-6 text-[#5B5A54]">Select your university, upload your DOCX, and review your formatting report.</p>
              <button onClick={openValidator} className="mt-7 bg-[#1F3A5F] px-7 py-3.5 text-sm font-semibold text-white transition hover:bg-[#172D49]">
                Check My Document →
              </button>
            </div>
          </section>
        </main>

        <footer className="border-t border-[#D8D4CB] bg-white">
          <div className="mx-auto flex max-w-6xl flex-col gap-3 px-5 py-6 text-sm text-[#8A887E] sm:flex-row sm:items-center sm:justify-between sm:px-8">
            <span>© 2026 Thesis Validator</span>
            <div className="flex flex-wrap items-center gap-4">
              <button onClick={openPrivacyPolicy} className="hover:text-[#1F3A5F] hover:underline">
                Privacy Policy
              </button>
              <span>Formatting assistance for academic documents</span>
            </div>
          </div>
        </footer>
      </div>
    )
  }

  return (
    <div id="validator" className="min-h-screen bg-[#F6F5F1] text-[#1E2A38]">
      <div className="mx-auto max-w-2xl px-4 py-10 sm:py-16">
        {/* Validator header */}
        <header className="mb-8">
          <div className="mb-5 flex items-center justify-between gap-4">
            <button onClick={goHome} className="text-sm text-[#1F3A5F] hover:underline">← Back to home</button>
          </div>
          <h1 className="font-serif text-3xl leading-tight sm:text-4xl">
            Thesis &amp; Lab Report Formatting Check
          </h1>
          <p className="mt-2 text-[#5B5A54]">
            Verify your document against your university's formatting requirements.
          </p>
        </header>

        {!result ? (
          <div className="border border-[#D8D4CB] bg-white">
            {/* Step 1: University */}
            <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
              <p className="text-xs font-medium tracking-wide text-[#8A887E]">Step 1</p>
              <label htmlFor="university" className="mt-1 block text-lg font-medium">
                University
              </label>
              <select
                id="university"
                value={university}
                onChange={handleUniversityChange}
                className="mt-4 w-full border border-[#D8D4CB] bg-white px-3 py-2.5 text-[#1E2A38] focus:outline-none focus:ring-2 focus:ring-[#1F3A5F]"
              >
                <option value="">Select a university...</option>
                {UNIVERSITIES.map((u) => (
                  <option key={u.name} value={u.name}>
                    {u.name}
                  </option>
                ))}
              </select>

              {/* Program dropdown — hidden until a university is chosen */}
              {selectedUniversity && (
                <div className="mt-4">
                  <label htmlFor="program" className="block text-sm font-medium text-[#1E2A38]">
                    Program / Document Type
                  </label>
                  <select
                    id="program"
                    value={profileName}
                    onChange={handleProfileChange}
                    className="mt-2 w-full border border-[#D8D4CB] bg-white px-3 py-2.5 text-[#1E2A38] focus:outline-none focus:ring-2 focus:ring-[#1F3A5F]"
                  >
                    <option value="">Select a program...</option>
                    {selectedUniversity.programs.map((program) => (
                      <option key={program.profile} value={program.profile}>
                        {program.label}
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>

            {/* Step 2: Upload */}
            <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
              <p className="text-xs font-medium tracking-wide text-[#8A887E]">Step 2</p>
              <p className="mt-1 text-lg font-medium">Document</p>

              {!selectedFile ? (
                <div
                  onClick={() => fileInputRef.current?.click()}
                  onDragOver={(e) => {
                    e.preventDefault()
                    setIsDragging(true)
                  }}
                  onDragLeave={() => setIsDragging(false)}
                  onDrop={handleDrop}
                  role="button"
                  tabIndex={0}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') fileInputRef.current?.click()
                  }}
                  className={`mt-4 flex cursor-pointer flex-col items-center justify-center border-2 border-dashed px-6 py-10 text-center transition-colors ${
                    isDragging
                      ? 'border-[#1F3A5F] bg-[#EEF2F6]'
                      : 'border-[#D8D4CB] hover:border-[#B7B2A4]'
                  }`}
                >
                  <p className="text-[#1E2A38]">
                    <span className="font-medium text-[#1F3A5F]">Click to browse</span>{' '}
                    or drag and drop your .docx file here
                  </p>
                  <p className="mt-1.5 text-sm text-[#8A887E]">Word documents (.docx) only · Maximum 10 MB</p>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".docx"
                    onChange={handleInputChange}
                    className="hidden"
                  />
                </div>
              ) : (
                <div className="mt-4 flex items-center justify-between border border-[#D8D4CB] bg-[#FAFAF8] px-4 py-3">
                  <span className="truncate text-sm font-medium text-[#1E2A38]">
                    {selectedFile.name}
                  </span>
                  <button
                    type="button"
                    onClick={handleRemoveFile}
                    className="ml-4 shrink-0 text-sm text-[#A8332B] hover:underline"
                  >
                    Remove
                  </button>
                </div>
              )}

              {fileError && <p className="mt-3 text-sm text-[#A8332B]">{fileError}</p>}
            </div>

            {/* Step 3: Validate */}
            <div className="p-6 sm:p-8">
              <p className="text-xs font-medium tracking-wide text-[#8A887E]">Step 3</p>
              <p className="mt-1 mb-4 text-lg font-medium">Run the check</p>

              <button
                onClick={handleValidate}
                disabled={!canValidate}
                className="w-full bg-[#1F3A5F] py-3 font-medium text-white transition-colors hover:bg-[#182E4A] disabled:cursor-not-allowed disabled:bg-[#C7C4BA]"
              >
                {isLoading ? 'Analyzing your document...' : 'Validate Document'}
              </button>

              <p className="mt-3 text-center text-xs leading-5 text-[#8A887E]">
                Your DOCX is processed temporarily for validation and is removed after processing.
                The current MVP does not intentionally store uploaded documents.
              </p>

              {requestError && (
                <p className="mt-3 text-center text-sm text-[#A8332B]">{requestError}</p>
              )}
            </div>
          </div>
        ) : (
          <div className="border border-[#D8D4CB] bg-white">
            {/* Result summary banner */}
            <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
              <div className="flex flex-col gap-6 sm:flex-row sm:items-start sm:justify-between">
                {/* Score block — the visual focal point of the summary */}
                {score !== null && (
                  <div className="shrink-0">
                    <p
                      className="font-serif text-5xl leading-none sm:text-6xl"
                      style={{ color: scoreColor(score) }}
                    >
                      {score}
                      <span className="text-xl text-[#8A887E] sm:text-2xl"> / 100</span>
                    </p>
                    {scoreLabel && (
                      <p
                        className="mt-1.5 text-base font-semibold"
                        style={{ color: scoreColor(score) }}
                      >
                        {scoreLabel}
                      </p>
                    )}
                    <div className="mt-3 h-1.5 w-full min-w-[180px] bg-[#EDEBE4] sm:w-48">
                      <div
                        className="h-full transition-all"
                        style={{ width: `${score}%`, backgroundColor: scoreColor(score) }}
                      />
                    </div>
                  </div>
                )}

                {/* Status + explanation + counts */}
                <div className="flex-1">
                  <div className="flex items-center gap-3">
                    <span
                      className={`inline-block h-2.5 w-2.5 rounded-full ${
                        result.is_valid ? 'bg-[#2F6846]' : 'bg-[#A8332B]'
                      }`}
                      aria-hidden="true"
                    />
                    <h2 className="font-serif text-2xl">
                      {result.is_valid ? 'Passed' : 'Needs Changes'}
                    </h2>
                  </div>

                  {score !== null && (
                    <p className="mt-2 text-[#5B5A54]">{scoreExplanation(score)}</p>
                  )}

                  <div className="mt-4 flex flex-wrap gap-x-6 gap-y-2 text-sm">
                    <p>
                      <span className="font-semibold text-[#1E2A38]">{blockingCount}</span>{' '}
                      <span className="text-[#5B5A54]">
                        blocking {blockingCount === 1 ? 'issue' : 'issues'} to fix
                      </span>
                    </p>
                    <p>
                      <span className="font-semibold text-[#1E2A38]">{infoCount}</span>{' '}
                      <span className="text-[#5B5A54]">
                        informational {infoCount === 1 ? 'note' : 'notes'}
                        {infoCount > 0 ? " (doesn't affect score)" : ''}
                      </span>
                    </p>
                  </div>
                </div>
              </div>
            </div>

            {score !== null && (
              <div className="border-b border-[#D8D4CB] bg-[#FAFAF8] px-6 py-5 sm:px-8">
                <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                  <div>
                    <p className="text-xs font-semibold uppercase tracking-wide text-[#8A887E]">
                      Score breakdown
                    </p>
                    <p className="mt-1 text-sm text-[#5B5A54]">
                      {blockingCount === 0
                        ? 'No blocking formatting violations were found.'
                        : `${blockingCount} blocking ${blockingCount === 1 ? 'issue' : 'issues'} were found.`}
                    </p>
                  </div>
                  <div className="flex flex-wrap gap-2 text-xs">
                    <span className="border border-[#CFE3D6] bg-[#F1F7F3] px-2.5 py-1 text-[#2F6846]">
                      {passedCheckCount} passed
                    </span>
                    <span className="border border-[#E8C9C5] bg-[#FBF0EE] px-2.5 py-1 text-[#A8332B]">
                      {blockingCount} issues
                    </span>
                    <span className="border border-[#E4D9B0] bg-[#FBF6E8] px-2.5 py-1 text-[#8A6D1F]">
                      {infoCount} info
                    </span>
                  </div>
                </div>
              </div>
            )}

            {/* Profile context — only shown fields the API actually returned */}
            {profileFields.length > 0 && (
              <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
                <h3 className="text-sm font-semibold text-[#1E2A38]">Checked Against</h3>
                <dl className="mt-3 grid grid-cols-1 gap-4 sm:grid-cols-3">
                  {profileFields.map((field) => (
                    <div key={field.label}>
                      <dt className="text-xs text-[#8A887E]">{field.label}</dt>
                      <dd className="mt-0.5 text-sm font-medium">{field.value}</dd>
                    </div>
                  ))}
                </dl>
              </div>
            )}

            {/* Check Summary */}
            <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
              <h3 className="text-sm font-semibold text-[#1E2A38]">Check Summary</h3>

              {checkSummary.length > 0 ? (
                <ul className="mt-3 grid grid-cols-1 gap-2 sm:grid-cols-2">
                  {checkSummary.map(({ category, status, reason }) => {
                    const isPassed = status === 'passed'
                    const isNotApplicable = status === 'not_applicable'

                    return (
                      <li
                        key={category}
                        className={`flex items-start gap-2 border px-3 py-2.5 text-sm ${
                          isPassed
                            ? 'border-[#CFE3D6] bg-[#F1F7F3]'
                            : isNotApplicable
                              ? 'border-[#DDD9D0] bg-[#F7F5F0]'
                              : 'border-[#E7CACA] bg-[#FBF2F2]'
                        }`}
                      >
                        <span
                          className={`mt-0.5 font-semibold ${
                            isPassed
                              ? 'text-[#2F6846]'
                              : isNotApplicable
                                ? 'text-[#8A887E]'
                                : 'text-[#9B3D3D]'
                          }`}
                          aria-hidden="true"
                        >
                          {isPassed ? '✓' : isNotApplicable ? '—' : '✕'}
                        </span>

                        <span className="min-w-0">
                          <span className="block font-medium text-[#1E2A38]">
                            {CATEGORY_LABELS[category] || category}
                          </span>
                          <span className="mt-0.5 block text-xs text-[#6F6B62]">
                            {isPassed
                              ? 'Passed'
                              : isNotApplicable
                                ? reason || 'Not enforced for this profile.'
                                : 'Needs attention'}
                          </span>
                        </span>
                      </li>
                    )
                  })}
                </ul>
              ) : (
                <p className="mt-2 text-sm text-[#8A887E]">
                  No check summary is available for this result.
                </p>
              )}
            </div>

            {(blockingCount > 0 || infoCount > 0) && (
              <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
                <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
                  <div>
                    <h3 className="text-sm font-semibold text-[#1E2A38]">Review Results</h3>
                    <p className="mt-1 text-sm text-[#8A887E]">
                      Filter the results to focus on what you need to fix.
                    </p>
                  </div>

                  <div className="grid grid-cols-1 gap-3 sm:grid-cols-2">
                    <label className="text-xs text-[#8A887E]">
                      <span className="mb-1 block">Show</span>
                      <select
                        value={issueFilter}
                        onChange={(e) => setIssueFilter(e.target.value)}
                        className="w-full border border-[#D8D4CB] bg-white px-3 py-2 text-sm text-[#1E2A38] focus:outline-none focus:ring-2 focus:ring-[#1F3A5F]"
                      >
                        <option value="all">All results</option>
                        <option value="errors">Errors only ({blockingCount})</option>
                        <option value="info">Info only ({infoCount})</option>
                      </select>
                    </label>

                    <label className="text-xs text-[#8A887E]">
                      <span className="mb-1 block">Category</span>
                      <select
                        value={categoryFilter}
                        onChange={(e) => setCategoryFilter(e.target.value)}
                        className="w-full border border-[#D8D4CB] bg-white px-3 py-2 text-sm text-[#1E2A38] focus:outline-none focus:ring-2 focus:ring-[#1F3A5F]"
                      >
                        <option value="all">All categories</option>
                        {issueCategories.map((category) => (
                          <option key={category} value={category}>
                            {getCategoryLabel(category)}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                </div>
              </div>
            )}

            {blockingCount > 0 && (
              <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
                <h3 className="text-sm font-semibold text-[#1E2A38]">
                  Formatting Issues ({visibleBlockingCount})
                </h3>

                {visibleBlockingCount > 0 ? (
                  <div className="mt-3 space-y-6">
                    {blockingByCategory.map(({ category, issues }) => (
                      <div key={category}>
                        <div className="flex items-baseline justify-between">
                          <h4 className="text-sm font-medium text-[#1E2A38]">
                            {getCategoryLabel(category)}
                          </h4>
                          <span className="text-xs text-[#8A887E]">
                            {issues.length} {issues.length === 1 ? 'issue group' : 'issue groups'}
                          </span>
                        </div>
                        <ul className="mt-2 divide-y divide-[#EDEBE4]">
                          {issues.map((issue, index) => (
                            <IssueCard key={`${issueKey(issue)}-${index}`} issue={issue} />
                          ))}
                        </ul>
                      </div>
                    ))}
                  </div>
                ) : (
                  <p className="mt-3 text-sm text-[#8A887E]">
                    No formatting issues match the selected filters.
                  </p>
                )}
              </div>
            )}

            {infoCount > 0 && (
              <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
                <h3 className="text-sm font-semibold text-[#1E2A38]">
                  Informational Notes ({visibleInfoCount})
                </h3>
                <p className="mt-1 text-sm text-[#8A887E]">
                  These do not mean your document failed. They can't be reliably verified from the
                  document's structure, are not counted as formatting errors, and do not reduce
                  your score.
                </p>
                {visibleInfoCount > 0 ? (
                  <ul className="mt-2 divide-y divide-[#EDEBE4]">
                    {groupedInfoIssues.map((issue, index) => (
                      <IssueCard key={`${issueKey(issue)}-info-${index}`} issue={issue} />
                    ))}
                  </ul>
                ) : (
                  <p className="mt-3 text-sm text-[#8A887E]">
                    No informational notes match the selected filters.
                  </p>
                )}
              </div>
            )}

            {(blockingCount > 0 || infoCount > 0) && !hasVisibleIssues && (
              <div className="border-b border-[#D8D4CB] p-6 sm:p-8">
                <p className="text-sm text-[#8A887E]">
                  Nothing matches these filters. Try selecting “All results” or “All categories.”
                </p>
              </div>
            )}

            {/* Trust / scope note */}
            <div className="border-b border-[#D8D4CB] bg-[#FAFAF8] px-6 py-4 sm:px-8">
              <p className="text-xs leading-relaxed text-[#8A887E]">
                This check evaluates the formatting rules configured for the selected
                university/program. It does not verify citations, grammar, or every visual aspect
                of a Word document. Pagination checks confirm structural facts in the document
                file — they do not verify exact rendered or printed page numbers.
              </p>
            </div>

            <div className="p-6 sm:p-8">
              <button
                onClick={handleReset}
                className="w-full border border-[#1F3A5F] py-3 font-medium text-[#1F3A5F] transition-colors hover:bg-[#EEF2F6]"
              >
                Validate Another Document
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}

export default App