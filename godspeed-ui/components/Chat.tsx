"use client"

import { useState, useRef, useEffect } from "react"
import axios from "axios"
import { Paperclip, Send, X } from "lucide-react"

const SMARTSHEET_LINK_REGEX = /https:\/\/app\.smartsheet\.com\/sheets\/[A-Za-z0-9]+(\?[^ ]*)?/

type Attachment = { id: string; name: string; mime_type: string }
type AttachmentGroup = { rowId: string; files: Attachment[] }
type SmartsheetPickerMessage = {
  type: "smartsheet-picker"
  attachments: AttachmentGroup[]
  selected: Set<string>
  sheet_id: string
  last_query?: string // for reselect tracking
}
type Message =
  | { role: "user" | "assistant"; content: string; smartsheet?: boolean }
  | SmartsheetPickerMessage

function addUniqueFiles(prev: File[], newFiles: File[]) {
  const existing = new Set(prev.map(f => `${f.name}-${f.size}`))
  return [
    ...prev,
    ...newFiles.filter(f => !existing.has(`${f.name}-${f.size}`))
  ]
}

export default function Chat() {
  const [messages, setMessages] = useState<Message[]>([])
  const [input, setInput] = useState("")
  const [files, setFiles] = useState<File[]>([])
  const [isDragging, setIsDragging] = useState(false)
  const [error, setError] = useState("")
  const [showExportMenu, setShowExportMenu] = useState<number | null>(null)
  const [showUploadMenu, setShowUploadMenu] = useState<number | null>(null)
  const [smartsheetRows, setSmartsheetRows] = useState<{ id: string, name: string }[]>([])
  const [lastExportedBlob, setLastExportedBlob] = useState<Blob | null>(null)
  const [lastExportedType, setLastExportedType] = useState<"pdf" | "docx" | null>(null)
  const scrollRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => { inputRef.current?.focus() }, [])
  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: "smooth" })
  }, [messages])

  // Drag-and-drop events
  useEffect(() => {
    const onDragOver = (e: DragEvent) => { e.preventDefault(); setIsDragging(true) }
    const onDrop = (e: DragEvent) => {
      e.preventDefault()
      setIsDragging(false)
      if (e.dataTransfer?.files?.length) {
        const dropped = Array.from(e.dataTransfer.files)
          .filter(f => [".pdf", ".docx", ".xlsx", ".txt"].some(ext => f.name.endsWith(ext)))
        setFiles(prev => addUniqueFiles(prev, dropped))
      }
    }
    const onDragLeave = (e: DragEvent) => { e.preventDefault(); setIsDragging(false) }
    window.addEventListener("dragover", onDragOver)
    window.addEventListener("drop", onDrop)
    window.addEventListener("dragleave", onDragLeave)
    return () => {
      window.removeEventListener("dragover", onDragOver)
      window.removeEventListener("drop", onDrop)
      window.removeEventListener("dragleave", onDragLeave)
    }
  }, [])

  // Helper: update Smartsheet picker selections
  function updatePickerSelections(selected: Set<string>) {
    setMessages(msgs => msgs.map(msg =>
      ("type" in msg && msg.type === "smartsheet-picker")
        ? { ...msg, selected: new Set(selected) }
        : msg
    ))
  }

  // PATCH: Structured export flow (first call /estimate/generate, then /estimate/export)
  const handleStructuredExport = async (content: string, format: "pdf" | "docx") => {
    const code = localStorage.getItem("access_code");
    try {
      // Step 1: Generate structured estimate JSON
      const genRes = await axios.post(
        "http://localhost:5001/estimate/generate",
        { content },
        { headers: { "X-Internal-Access": code } }
      );
      const estimateData = genRes.data.estimate;
      // Step 2: Export that as PDF/DOCX
      const expRes = await axios.post(
        "http://localhost:5001/estimate/export",
        { estimate_data: estimateData, format },
        { headers: { "X-Internal-Access": code }, responseType: "blob" }
      );
      const blob = new Blob(
        [expRes.data],
        {
          type:
            format === "pdf"
              ? "application/pdf"
              : "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        }
      );
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = `estimate.${format}`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
      setLastExportedBlob(blob);
      setLastExportedType(format);
    } catch (e) {
      setError("Export failed. Try again.");
    }
  };

  // Fetch Smartsheet rows for upload
  const handleShowUpload = async (sheet_id: string) => {
    const code = localStorage.getItem("access_code")
    try {
      const res = await axios.post(
        "http://localhost:5001/smartsheet/list-rows",
        { sheet_id },
        { headers: { "X-Internal-Access": code } }
      )
      setSmartsheetRows(res.data.rows)
    } catch {
      setError("Failed to fetch rows.")
    }
  }

  // Upload exported file as Smartsheet attachment
  const handleUploadToSmartsheet = async (row_id: string, filename?: string) => {
    if (!lastExportedBlob || !lastExportedType) {
      setError("Please export a file first.")
      return
    }
    const code = localStorage.getItem("access_code")
    try {
      const base64 = await new Promise<string>((resolve, reject) => {
        const reader = new FileReader()
        reader.onload = () => {
          const base64str = reader.result?.toString().split(',')[1]
          if (base64str) resolve(base64str)
          else reject("Encoding failed")
        }
        reader.onerror = reject
        reader.readAsDataURL(lastExportedBlob)
      })
      await axios.post(
        "http://localhost:5001/smartsheet/upload",
        { file_buffer: base64, row_id, filename: filename || `estimate.${lastExportedType}` },
        { headers: { "X-Internal-Access": code } }
      )
      alert("Uploaded to Smartsheet!")
      setShowUploadMenu(null)
    } catch {
      setError("Upload failed. Try again.")
    }
  }

  // Show Smartsheet picker by pushing last used picker message
  function showSmartsheetPickerAgain() {
    const pickerMsg = messages.findLast(m => "type" in m && m.type === "smartsheet-picker") as SmartsheetPickerMessage | undefined
    if (pickerMsg) {
      setMessages(msgs => [
        ...msgs,
        { ...pickerMsg, selected: new Set() }
      ])
      setInput("")
    }
  }

  // Main send handler
  const handleSend = async (e?: React.FormEvent) => {
    if (e) e.preventDefault()
    setError("")
    const code = localStorage.getItem("access_code")
    if (!code) {
      window.location.href = "/login"
      return
    }

    // 1️⃣ If showing Smartsheet picker, analyze selected attachments
    const lastMsg = messages[messages.length - 1]
    if (
      lastMsg &&
      "type" in lastMsg &&
      lastMsg.type === "smartsheet-picker" &&
      lastMsg.selected &&
      lastMsg.selected.size > 0
    ) {
      // Preview: which files are being analyzed?
      const picker = lastMsg as SmartsheetPickerMessage
      const selectedFiles: string[] = []
      picker.attachments.forEach(group => {
        group.files.forEach(file => {
          if (picker.selected.has(String(file.id))) selectedFiles.push(file.name)
        })
      })

      setMessages(msgs => [
        ...msgs,
        { role: "user", content: input || `Analyze selected Smartsheet files: ${selectedFiles.join(", ")}` }
      ])
      setInput("")
      try {
        const res = await axios.post(
          "http://localhost:5001/smartsheet/analyze-batch",
          {
            sheet_id: picker.sheet_id,
            attachment_ids: Array.from(picker.selected),
            query: picker.last_query || input,
            additional_text: ""
          },
          { headers: { "X-Internal-Access": code } }
        )
        setMessages(msgs => [
          ...msgs,
          {
            role: "assistant",
            content: res.data.response || JSON.stringify(res.data),
            smartsheet: true // mark as smartsheet output
          }
        ])
      } catch {
        setMessages(msgs => [
          ...msgs,
          { role: "assistant", content: "⚠️ Error analyzing attachments.", smartsheet: true }
        ])
        setError("Server error. Please try again.")
      }
      return
    }

    // 2️⃣ Detect Smartsheet link and show picker
    if (SMARTSHEET_LINK_REGEX.test(input.trim())) {
      const link = input.trim()
      setMessages(msgs => [
        ...msgs,
        { role: "user", content: link },
        { role: "assistant", content: "Fetching attachments from Smartsheet…" }
      ])
      setInput("")
      try {
        const res = await axios.post(
          "http://localhost:5001/smartsheet/list-attachments",
          { sheet_url: link },
          { headers: { "X-Internal-Access": code } }
        )
        const attachments: AttachmentGroup[] = Object.entries(res.data.attachments).map(
          ([rowId, files]: any) => ({ rowId, files })
        )
        setMessages(msgs => [
          ...msgs.slice(0, -1),
          {
            type: "smartsheet-picker",
            attachments,
            selected: new Set(),
            sheet_id: res.data.sheet_id,
            last_query: input
          }
        ])
      } catch {
        setMessages(msgs => [
          ...msgs.slice(0, -1),
          { role: "assistant", content: "❌ Failed to load attachments." }
        ])
      }
      return
    }

    // 3️⃣ Standard file upload or chat
    if (!input.trim() && files.length === 0) return
    setMessages(msgs => [
      ...msgs,
      { role: "user", content: input || (files.map(f => f.name).join(", ")) }
    ])
    try {
      let response
      if (files.length > 0) {
        const form = new FormData()
        files.forEach(file => form.append("files", file))
        form.append("query", input)
        response = await axios.post(
          "http://localhost:5001/chat/upload-analyze",
          form,
          { headers: { "X-Internal-Access": code } }
        )
      } else {
        response = await axios.post(
          "http://localhost:5001/chat/analyze",
          { query: input, content: "" },
          { headers: { "X-Internal-Access": code } }
        )
      }
      setMessages(msgs => [
        ...msgs,
        { role: "assistant", content: response.data.response || JSON.stringify(response.data) }
      ])
      setInput("")
      setFiles([])
    } catch {
      setMessages(msgs => [
        ...msgs,
        { role: "assistant", content: "⚠️ Server error. Please try again." }
      ])
      setError("Server error. Please try again.")
    }
  }

  // Render UI
  return (
    <div
      className={`h-screen flex flex-col ${isDragging ? "bg-blue-50" : ""}`}
      onDragOver={e => { e.preventDefault(); setIsDragging(true) }}
      onDrop={e => {
        e.preventDefault()
        setIsDragging(false)
        if (e.dataTransfer?.files?.length) {
          const dropped = Array.from(e.dataTransfer.files)
            .filter(f => [".pdf", ".docx", ".xlsx", ".txt"].some(ext => f.name.endsWith(ext)))
          setFiles(prev => addUniqueFiles(prev, dropped))
        }
      }}
      onDragLeave={e => { e.preventDefault(); setIsDragging(false) }}
    >
      {/* Chat history */}
      <div ref={scrollRef} className="flex-1 overflow-y-auto px-4 py-8 bg-background">
        {messages.length === 0 && (
          <div className="text-center text-muted-foreground pt-20">
            Upload files, paste a Smartsheet link, or ask a question to begin.
          </div>
        )}
        {messages.map((msg, i) => {
          // Smartsheet attachment picker
          if ("type" in msg && msg.type === "smartsheet-picker") {
            const selected = msg.selected ?? new Set()
            const allFiles: Attachment[] = msg.attachments.flatMap(g => g.files)
            const allIds = allFiles.map(f => String(f.id))
            const isAllSelected = allIds.every(id => selected.has(id))
            return (
              <div key={i} className="mb-6">
                <div className="font-semibold mb-2">Select attachments to analyze:</div>
                <div className="mb-2 flex flex-col gap-1 max-h-64 overflow-y-auto border rounded p-2 bg-muted">
                  <label className="flex items-center gap-2 font-medium">
                    <input
                      type="checkbox"
                      checked={isAllSelected}
                      onChange={e => {
                        updatePickerSelections(
                          e.target.checked ? new Set(allIds) : new Set()
                        )
                      }}
                    />
                    Select all
                  </label>
                  {msg.attachments.map(group => (
                    <div key={group.rowId}>
                      <div className="text-xs font-semibold text-muted-foreground mb-1">
                        Row {group.rowId}
                      </div>
                      {group.files.map(file => (
                        <label key={file.id} className="flex items-center gap-2 ml-3">
                          <input
                            type="checkbox"
                            checked={selected.has(String(file.id))}
                            onChange={e => {
                              const next = new Set(selected)
                              if (e.target.checked) next.add(String(file.id))
                              else next.delete(String(file.id))
                              updatePickerSelections(next)
                            }}
                          />
                          <span className="text-xs">{file.name}</span>
                          <span className="text-muted-foreground text-xs ml-1">{file.mime_type}</span>
                        </label>
                      ))}
                    </div>
                  ))}
                </div>
                <div className="text-xs text-muted-foreground">
                  (You can add optional notes below before sending.)
                </div>
              </div>
            )
          }
          // Standard chat/assistant message (includes export/upload, with smartsheet label and reselect)
          const isAssistant = "role" in msg && msg.role === "assistant"
          // Did this message come after a Smartsheet picker?
          const isFromSmartsheet =
            msg && isAssistant && msg.smartsheet

          return (
            <div key={i} className={`mb-6 flex ${isAssistant ? "justify-start" : "justify-end"}`}>
              <div className={`max-w-lg px-4 py-2 rounded-xl whitespace-pre-wrap ${isAssistant ? "bg-muted" : "bg-primary text-white"}`}>
                {msg.content}
                {isFromSmartsheet && (
                  <div className="text-xs font-semibold mt-1 text-blue-700">Smartsheet Output</div>
                )}
                {isAssistant && (
                  <div className="flex gap-2 mt-2">
                    <button
                      className="px-2 py-1 bg-gray-200 rounded text-xs"
                      onClick={() => setShowExportMenu(i)}
                    >
                      Export
                    </button>
                    {isFromSmartsheet && (
                      <button
                        className="px-2 py-1 bg-blue-200 rounded text-xs"
                        onClick={async () => {
                          if (!lastExportedBlob) {
                            alert("Export first, then upload to Smartsheet.")
                            return
                          }
                          // Use last Smartsheet picker for row selection
                          const pickerMsg = messages.findLast(
                            m => "type" in m && m.type === "smartsheet-picker"
                          ) as SmartsheetPickerMessage | undefined
                          if (pickerMsg?.sheet_id) {
                            await handleShowUpload(pickerMsg.sheet_id)
                            setShowUploadMenu(i)
                          }
                        }}
                      >
                        Upload to Smartsheet
                      </button>
                    )}
                    {isFromSmartsheet && (
                      <button
                        className="px-2 py-1 bg-yellow-100 rounded text-xs"
                        onClick={showSmartsheetPickerAgain}
                      >
                        Reselect Smartsheet attachments
                      </button>
                    )}
                  </div>
                )}
                {/* Export menu popup */}
                {showExportMenu === i && (
                  <div className="mt-1 flex gap-2">
                    <button
                      className="px-2 py-1 bg-gray-300 rounded text-xs"
                      onClick={async () => {
                        await handleStructuredExport(msg.content, "pdf")
                        setShowExportMenu(null)
                      }}
                    >PDF</button>
                    <button
                      className="px-2 py-1 bg-gray-300 rounded text-xs"
                      onClick={async () => {
                        await handleStructuredExport(msg.content, "docx")
                        setShowExportMenu(null)
                      }}
                    >DOCX</button>
                    <button className="text-xs" onClick={() => setShowExportMenu(null)}>Cancel</button>
                  </div>
                )}

                {/* Upload row picker */}
                {showUploadMenu === i && (
                  <div className="mt-1">
                    <div className="text-xs mb-1">Pick Smartsheet Row:</div>
                    <select
                      className="border px-2 py-1 rounded text-xs"
                      onChange={e =>
                        handleUploadToSmartsheet(e.target.value, `estimate.${lastExportedType}`)
                      }
                    >
                      <option value="">Select row...</option>
                      {smartsheetRows.map(row => (
                        <option key={row.id} value={row.id}>{row.name}</option>
                      ))}
                    </select>
                    <button className="text-xs ml-2" onClick={() => setShowUploadMenu(null)}>Cancel</button>
                  </div>
                )}
              </div>
            </div>
          )
        })}
      </div>

      {/* Input dock */}
      <form
        className="flex gap-2 p-4 border-t bg-background items-end"
        onSubmit={handleSend}
      >
        {/* Paperclip (file picker) */}
        <label className="flex items-center cursor-pointer">
          <Paperclip className="w-5 h-5 mr-2 text-muted-foreground" />
          <input
            type="file"
            accept=".pdf,.docx,.xlsx,.txt"
            className="hidden"
            multiple
            onChange={e => {
              if (e.target.files) {
                setFiles(prev => addUniqueFiles(prev, Array.from(e.target.files)))
                e.target.value = "" // reset picker so same file can be reselected after delete
              }
            }}
          />
        </label>
        {/* File preview chips */}
        {files.length > 0 && (
          <div className="flex flex-wrap gap-1">
            {files.map((f, idx) => (
              <div key={idx} className="flex items-center gap-1 bg-muted px-2 py-1 rounded">
                <span className="text-xs">{f.name}</span>
                <button
                  type="button"
                  onClick={() => setFiles(fs => fs.filter((_, i) => i !== idx))}
                  className="text-red-500 text-xs"
                  aria-label="Remove file"
                >
                  <X className="w-3 h-3" />
                </button>
              </div>
            ))}
          </div>
        )}
        <input
          ref={inputRef}
          className="flex-1 border p-2 rounded"
          value={input}
          onChange={e => setInput(e.target.value)}
          placeholder="Type a message, drop files, or paste a Smartsheet link…"
        />
        <button
          type="submit"
          className="bg-black text-white px-4 py-2 rounded"
          disabled={
            messages.length > 0 &&
              "type" in messages[messages.length - 1] &&
              messages[messages.length - 1].type === "smartsheet-picker"
              ? !(messages[messages.length - 1] as any).selected || !(messages[messages.length - 1] as any).selected.size
              : (!input.trim() && files.length === 0)
          }
        >
          <Send className="w-5 h-5" />
        </button>
      </form>
      {error && (
        <div className="text-red-500 text-xs text-center pb-2">{error}</div>
      )}
      {/* Drag overlay visual */}
      {isDragging && (
        <div className="fixed inset-0 bg-blue-200 bg-opacity-30 flex items-center justify-center pointer-events-none z-50">
          <div className="p-6 bg-white rounded-xl shadow text-xl font-semibold border-2 border-blue-500 border-dashed">
            Drop your files here
          </div>
        </div>
      )}
    </div>
  )
}
