"use client"

import type React from "react"

import { useState, useRef } from "react"
import { Button } from "../components/ui/button"
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../components/ui/card"
import { Tabs, TabsContent, TabsList, TabsTrigger } from "../components/ui/tabs"
import { Loader2 } from "lucide-react"
import { Download } from "lucide-react"
import Papa from "papaparse"
import { Alert, AlertDescription } from "../components/ui/alert"
import { StatsTableWithDownload } from "@/src/components/custom/stats-table-whith-download"
import { processCSVData, generateExcel } from "@/src/lib/data-utils"
// Importar el icono personalizado
import { UploadIcon } from "@/src/components/custom/upload-icon"

interface Participant {
  genero: string
  facultad: string
  programaAcademico: string
  grupoCultural?: string
}

interface CrossTabStats {
  [key: string]: {
    [gender: string]: number
    total: number
  }
}

export default function Home() {
  const [loading, setLoading] = useState(false)
  const [participants, setParticipants] = useState<Participant[]>([])
  const [statsByFaculty, setStatsByFaculty] = useState<CrossTabStats>({})
  const [statsByProgram, setStatsByProgram] = useState<CrossTabStats>({})
  const [statsByGrupoCultural, setStatsByGrupoCultural] = useState<CrossTabStats>({})
  const [statsByGrupoIndividual, setStatsByGrupoIndividual] = useState<{
    [grupo: string]: { byFaculty: CrossTabStats; byProgram: CrossTabStats }
  }>({})
  const [uniqueGenders, setUniqueGenders] = useState<string[]>([])
  const [error, setError] = useState<string | null>(null)
  const [success, setSuccess] = useState<string | null>(null)
  const [fileName, setFileName] = useState<string | null>(null)

  const [dataType, setDataType] = useState<"baile" | "cultural">("baile")
  const fileInputRef = useRef<HTMLInputElement>(null)
  const culturalFileInputRef = useRef<HTMLInputElement>(null)

  const handleFileUpload = (event: React.ChangeEvent<HTMLInputElement>, type: "baile" | "cultural") => {
    const file = event.target.files?.[0]
    if (!file) return

    setLoading(true)
    setError(null)
    setSuccess(null)
    setFileName(file.name)

    // Usar FileReader para leer el archivo como texto
    const reader = new FileReader()
    reader.onload = (e) => {
      const csvText = e.target?.result as string
      if (!csvText) {
        setError("Error al leer el archivo CSV.")
        setLoading(false)
        return
      }

      try {
        console.log("Procesando archivo CSV:", file.name)
        console.log("Tipo de datos:", type)

        const results = Papa.parse(csvText, { header: true })
        console.log("Datos parseados:", results.data.slice(0, 2)) // Mostrar las primeras 2 filas para debug

        const {
          participants,
          statsByFaculty,
          statsByProgram,
          statsByGrupoCultural,
          statsByGrupoIndividual,
          uniqueGenders,
        } = processCSVData(results.data as any[], type)

        console.log("Participantes procesados:", participants.length)
        console.log("Géneros únicos:", uniqueGenders)

        setParticipants(participants)
        setStatsByFaculty(statsByFaculty)
        setStatsByProgram(statsByProgram)
        setStatsByGrupoCultural(statsByGrupoCultural)
        setStatsByGrupoIndividual(statsByGrupoIndividual)
        setUniqueGenders(uniqueGenders)
        setDataType(type)

        setSuccess(
          `Archivo "${file.name}" procesado correctamente. Se encontraron ${participants.length} participantes.`,
        )
        setLoading(false)
      } catch (err) {
        console.error("Error al procesar archivo:", err)
        setError(`Error al procesar el archivo: ${err instanceof Error ? err.message : String(err)}`)
        setLoading(false)
      }
    }

    reader.onerror = () => {
      setError("Error al leer el archivo CSV.")
      setLoading(false)
    }

    reader.readAsText(file)
  }

  const downloadExcel = async () => {
    if (participants.length === 0) {
      setError("No hay datos para descargar. Por favor, carga un archivo CSV primero.")
      return
    }

    try {
      await generateExcel({
        statsByFaculty,
        statsByProgram,
        statsByGrupoCultural,
        statsByGrupoIndividual,
        uniqueGenders,
        dataType,
      })

      setSuccess("Archivo Excel generado correctamente.")
    } catch (err) {
      setError(`Error al generar el archivo Excel: ${err instanceof Error ? err.message : String(err)}`)
    }
  }

  const triggerFileInput = (type: "baile" | "cultural") => {
    if (type === "baile" && fileInputRef.current) {
      fileInputRef.current.click()
    } else if (type === "cultural" && culturalFileInputRef.current) {
      culturalFileInputRef.current.click()
    }
  }

  const clearData = () => {
    setParticipants([])
    setStatsByFaculty({})
    setStatsByProgram({})
    setStatsByGrupoCultural({})
    setStatsByGrupoIndividual({})
    setUniqueGenders([])
    setFileName(null)
    setError(null)
    setSuccess(null)
  }

  return (
    <div className="container mx-auto py-10">
      <h1 className="text-2xl font-bold text-red-600 mb-6 text-center">AREA DE CULTURA DE LA UNIVERSIDAD DEL VALLE</h1>
      <Card className="mb-8">
        <CardHeader>
          <CardTitle>Estadísticas de Asistencia - Universidad del Valle</CardTitle>
          <CardDescription>Análisis de participantes por facultad, programa académico y grupo cultural</CardDescription>
        </CardHeader>
        <CardContent>
          <Tabs defaultValue="baile" className="mb-6">
            <TabsList className="mb-4">
              <TabsTrigger value="baile">Baile Recreativo</TabsTrigger>
              <TabsTrigger value="cultural">Grupos Culturales</TabsTrigger>
            </TabsList>

            <TabsContent value="baile">
              <div className="p-4 border rounded-lg bg-muted/20">
                <h3 className="text-lg font-medium mb-2">Cargar archivo CSV - Baile Recreativo</h3>
                <p className="text-sm text-muted-foreground mb-4">
                  Sube un archivo CSV con el formato de respuestas del formulario de asistencia de baile recreativo.
                </p>

                <div className="flex flex-col sm:flex-row gap-3">
                  <Button
                    onClick={() => triggerFileInput("baile")}
                    className="flex items-center gap-2"
                    variant="outline"
                  >
                    <UploadIcon />
                    Seleccionar archivo CSV
                  </Button>

                  {fileName && (
                    <Button onClick={clearData} variant="destructive" className="flex items-center gap-2">
                      Limpiar datos
                    </Button>
                  )}

                  <input
                    type="file"
                    ref={fileInputRef}
                    onChange={(e) => handleFileUpload(e, "baile")}
                    accept=".csv"
                    className="hidden"
                  />
                </div>
              </div>
            </TabsContent>

            <TabsContent value="cultural">
              <div className="p-4 border rounded-lg bg-muted/20">
                <h3 className="text-lg font-medium mb-2">Cargar archivo CSV - Grupos Culturales</h3>
                <p className="text-sm text-muted-foreground mb-4">
                  Sube un archivo CSV con el formato de respuestas del formulario de asistencia de grupos culturales.
                </p>

                <div className="flex flex-col sm:flex-row gap-3">
                  <Button
                    onClick={() => triggerFileInput("cultural")}
                    className="flex items-center gap-2"
                    variant="outline"
                  >
                    <UploadIcon />
                    Seleccionar archivo CSV
                  </Button>

                  {fileName && (
                    <Button onClick={clearData} variant="destructive" className="flex items-center gap-2">
                      Limpiar datos
                    </Button>
                  )}

                  <input
                    type="file"
                    ref={culturalFileInputRef}
                    onChange={(e) => handleFileUpload(e, "cultural")}
                    accept=".csv"
                    className="hidden"
                  />
                </div>
              </div>
            </TabsContent>
          </Tabs>

          {fileName && (
            <div className="mt-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
              <p className="text-sm">
                <span className="font-medium">Archivo cargado:</span> {fileName}
              </p>
              <p className="text-sm text-muted-foreground">
                <span className="font-medium">Participantes encontrados:</span> {participants.length}
              </p>
              {uniqueGenders.length > 0 && (
                <p className="text-sm text-muted-foreground">
                  <span className="font-medium">Géneros identificados:</span> {uniqueGenders.join(", ")}
                </p>
              )}
            </div>
          )}

          {error && (
            <Alert variant="destructive" className="mb-4">
              <AlertDescription>{error}</AlertDescription>
            </Alert>
          )}

          {success && (
            <Alert className="mb-4 border-green-500 text-green-700 bg-green-50">
              <AlertDescription>{success}</AlertDescription>
            </Alert>
          )}

          {loading ? (
            <div className="flex justify-center items-center h-40">
              <Loader2 className="h-8 w-8 animate-spin text-primary" />
              <span className="ml-2">Procesando datos del archivo seleccionado...</span>
            </div>
          ) : participants.length > 0 ? (
            <>
              <div className="flex justify-end mb-4">
                <Button onClick={downloadExcel} className="flex items-center gap-2">
                  <Download className="h-4 w-4" />
                  Descargar Todas las Estadísticas en Excel
                </Button>
              </div>

              <Tabs
                defaultValue={
                  dataType === "cultural" && Object.keys(statsByGrupoCultural).length > 0 ? "cultural" : "faculty"
                }
              >
                <TabsList className="mb-4">
                  <TabsTrigger value="faculty">Por Facultad</TabsTrigger>
                  <TabsTrigger value="program">Por Programa Académico</TabsTrigger>
                  {dataType === "cultural" && Object.keys(statsByGrupoCultural).length > 0 && (
                    <TabsTrigger value="cultural">Por Grupo Cultural</TabsTrigger>
                  )}
                  {dataType === "cultural" && Object.keys(statsByGrupoIndividual).length > 0 && (
                    <TabsTrigger value="individual">Estadísticas Individuales</TabsTrigger>
                  )}
                </TabsList>

                <TabsContent value="faculty">
                  <StatsTableWithDownload
                    title="Facultad"
                    stats={statsByFaculty}
                    genders={uniqueGenders}
                    fileName="Estadisticas_por_Facultad.xlsx"
                  />
                </TabsContent>

                <TabsContent value="program">
                  <StatsTableWithDownload
                    title="Programa Académico"
                    stats={statsByProgram}
                    genders={uniqueGenders}
                    fileName="Estadisticas_por_Programa.xlsx"
                  />
                </TabsContent>

                {dataType === "cultural" && Object.keys(statsByGrupoCultural).length > 0 && (
                  <TabsContent value="cultural">
                    <StatsTableWithDownload
                      title="Grupo Cultural"
                      stats={statsByGrupoCultural}
                      genders={uniqueGenders}
                      fileName="Estadisticas_por_Grupo_Cultural.xlsx"
                    />
                  </TabsContent>
                )}

                {dataType === "cultural" && Object.keys(statsByGrupoIndividual).length > 0 && (
                  <TabsContent value="individual">
                    <div className="space-y-8">
                      {Object.entries(statsByGrupoIndividual).map(([grupo, stats]) => (
                        <div key={grupo} className="border rounded-lg p-6">
                          <h2 className="text-xl font-bold mb-6 text-center text-blue-600">{grupo}</h2>

                          <div className="grid gap-8 md:grid-cols-1 lg:grid-cols-2">
                            <div>
                              <StatsTableWithDownload
                                title={`${grupo} - Por Facultad`}
                                stats={stats.byFaculty}
                                genders={uniqueGenders}
                                fileName={`${grupo}_por_Facultad.xlsx`}
                              />
                            </div>

                            <div>
                              <StatsTableWithDownload
                                title={`${grupo} - Por Programa`}
                                stats={stats.byProgram}
                                genders={uniqueGenders}
                                fileName={`${grupo}_por_Programa.xlsx`}
                              />
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  </TabsContent>
                )}
              </Tabs>
            </>
          ) : (
            <div className="text-center py-10 text-muted-foreground">
              <p>No hay datos para mostrar. Por favor, selecciona y carga un archivo CSV para generar estadísticas.</p>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  )
}
