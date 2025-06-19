"use client"

import type React from "react"
import { Button } from "@/src/components/ui/button"
import { Upload } from "lucide-react"

interface FileUploaderProps {
  title: string
  description: string
  onFileSelect: (event: React.ChangeEvent<HTMLInputElement>) => void
  onSampleLoad: () => void
  fileInputRef: React.RefObject<HTMLInputElement>
}

export function FileUploader({ title, description, onFileSelect, onSampleLoad, fileInputRef }: FileUploaderProps) {
  const handleButtonClick = () => {
    if (fileInputRef.current) {
      fileInputRef.current.click()
    }
  }

  return (
    <div className="p-4 border rounded-lg bg-muted/20">
      <h3 className="text-lg font-medium mb-2">{title}</h3>
      <p className="text-sm text-muted-foreground mb-4">{description}</p>

      <div className="flex flex-col sm:flex-row gap-3">
        <Button onClick={handleButtonClick} className="flex items-center gap-2" variant="outline">
          <Upload className="h-4 w-4" />
          Seleccionar archivo CSV
        </Button>

        <Button onClick={onSampleLoad} variant="secondary" className="flex items-center gap-2">
          Cargar datos de muestra
        </Button>

        <input type="file" ref={fileInputRef} onChange={onFileSelect} accept=".csv" className="hidden" />
      </div>
    </div>
  )
}
