import { useState, useRef, useEffect } from "react"
import axios from "axios"
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from "@/components/ui/card"
import { Input } from "@/components/ui/input"
import { Label } from "@/components/ui/label"
import { Button } from "@/components/ui/button"

export default function Login({ onLogin }: { onLogin: (token: string) => void }) {
  const [code, setCode] = useState("")
  const [error, setError] = useState("")
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    inputRef.current?.focus()
  }, [])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    try {
      const res = await axios.post("http://localhost:5001/auth", { code })
      if (res.data.status === "ok") {
        localStorage.setItem("access_code", code)
        onLogin(code)
      } else {
        setError("Invalid access code.")
      }
    } catch {
      setError("Server error. Please try again.")
    }
  }

  return (
    <div className="h-screen w-screen flex items-center justify-center bg-background px-4">
      {/* STEP 5.3 TEST BAR */}
      <div className="fixed top-0 left-0 w-full bg-red-300 text-center text-lg font-bold py-2 z-50">
        Tailwind is working!
      </div>

      <form onSubmit={handleSubmit} className="w-full max-w-sm">
        <Card className="w-full">
          <CardHeader>
            <CardTitle className="text-center">Godspeed Estimator</CardTitle>
          </CardHeader>

          <CardContent className="space-y-4">
            <div className="space-y-2">
              <Label htmlFor="code">Access Code</Label>
              <Input
                ref={inputRef}
                id="code"
                type="password"
                placeholder="••••••••"
                value={code}
                onChange={(e) => setCode(e.target.value)}
              />
              {error && <p className="text-sm text-red-500">{error}</p>}
            </div>
          </CardContent>

          <CardFooter>
            <Button type="submit" className="w-full">
              Enter
            </Button>
          </CardFooter>
        </Card>
      </form>
    </div>
  )
}
