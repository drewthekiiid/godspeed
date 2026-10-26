import { useState, useRef } from "react"
import axios from "axios"

interface ChatProps {
  accessCode: string
  onLogout: () => void
}

export default function Chat({ accessCode, onLogout }: ChatProps) {
  const [messages, setMessages] = useState<any[]>([])
  const [input, setInput] = useState("")
  const fileInputRef = useRef<HTMLInputElement>(null)

  const sendMessage = async () => {
    if (!input && !fileInputRef.current?.files?.length) return

    const userMsg = { role: "user", content: input || "[File Upload]" }
    setMessages((prev) => [...prev, userMsg, { role: "assistant", content: "..." }])

    try {
      let responseText = ""

      if (fileInputRef.current?.files?.length) {
        const file = fileInputRef.current.files[0]
        const form = new FormData()
        form.append("file", file)

        const res = await axios.post(
          "http://localhost:5001/chat/upload-analyze",
          form,
          {
            headers: { "X-Internal-Access": accessCode },
          }
        )
        responseText = res.data.response
      } else {
        const res = await axios.post(
          "http://localhost:5001/chat/analyze",
          {
            query: input,
            content: "",
          },
          {
            headers: { "X-Internal-Access": accessCode },
          }
        )
        responseText = res.data.response
      }

      setMessages((prev) => {
        const updated = [...prev]
        updated[updated.length - 1] = { role: "assistant", content: responseText }
        return updated
      })

      setInput("")
      if (fileInputRef.current) fileInputRef.current.value = ""
    } catch (err) {
      console.error(err)
    }
  }

  return (
    <div className="flex flex-col h-screen">
      <div className="flex justify-between items-center px-4 py-2 border-b bg-white">
        <h1 className="text-lg font-bold">Godspeed Estimator</h1>
        <button
          onClick={onLogout}
          className="text-sm text-gray-600 hover:text-black"
        >
          Logout
        </button>
      </div>

      <div className="flex-1 overflow-y-scroll px-4 py-6 space-y-4 bg-gray-50">
        {messages.map((msg, i) => (
          <div
            key={i}
            className={`p-3 rounded-md max-w-xl ${
              msg.role === "user"
                ? "bg-black text-white self-end ml-auto"
                : "bg-white border self-start"
            }`}
          >
            {msg.content}
          </div>
        ))}
      </div>

      <div className="p-4 border-t bg-white flex gap-2 items-center">
        <input
          type="text"
          placeholder="Type here..."
          value={input}
          onChange={(e) => setInput(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && sendMessage()}
          className="flex-1 border rounded px-3 py-2"
        />
        <input ref={fileInputRef} type="file" className="hidden" />
        <button
          onClick={() => fileInputRef.current?.click()}
          className="text-lg"
        >
          📎
        </button>
        <button
          onClick={sendMessage}
          className="bg-black text-white px-4 py-2 rounded"
        >
          Send
        </button>
      </div>
    </div>
  )
}