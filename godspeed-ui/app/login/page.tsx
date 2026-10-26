"use client"
import { useState, useRef, useEffect } from "react"
import { Card, CardHeader, CardContent, CardFooter, CardTitle } from "@/components/ui/card"
import { Label } from "@/components/ui/label"
import { Input } from "@/components/ui/input"
import { Button } from "@/components/ui/button"

export default function LoginPage() {
  const [code, setCode] = useState("")
  const [error, setError] = useState("")
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    inputRef.current?.focus()
  }, [])

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault()
    setError("")
    try {
      const res = await fetch("http://localhost:5001/auth", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ code }),
      })
      const data = await res.json()
      if (data.status === "ok") {
        localStorage.setItem("access_code", code)
        // TODO: Route to your app’s main page or dashboard
        window.location.href = "/"
      } else {
        setError("Invalid access code.")
      }
    } catch {
      setError("Server error. Please try again.")
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center bg-background">
      <form onSubmit={handleSubmit} className="w-full max-w-sm">
        <Card className="w-full">
          <CardHeader>
            <CardTitle className="text-center text-2xl font-bold">Godspeed Estimator</CardTitle>
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
                onChange={e => setCode(e.target.value)}
                autoComplete="off"
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