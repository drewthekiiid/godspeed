import { Button } from "@/components/ui/button"

export default function App() {
  return (
    <>
      <div className="bg-red-500 text-white p-4 text-center font-bold">
        TAILWIND IS WORKING
      </div>
      <div className="min-h-screen flex items-center justify-center bg-background text-foreground">
        <Button variant="default">Godspeed Online</Button>
      </div>
    </>
  )
}
