import ExcelJS from "exceljs"

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

interface ProcessedData {
  participants: Participant[]
  statsByFaculty: CrossTabStats
  statsByProgram: CrossTabStats
  statsByGrupoCultural: CrossTabStats
  statsByGrupoIndividual: { [grupo: string]: { byFaculty: CrossTabStats; byProgram: CrossTabStats } }
  uniqueGenders: string[]
}

interface ExcelGenerationParams {
  statsByFaculty: CrossTabStats
  statsByProgram: CrossTabStats
  statsByGrupoCultural: CrossTabStats
  statsByGrupoIndividual?: { [grupo: string]: { byFaculty: CrossTabStats; byProgram: CrossTabStats } }
  uniqueGenders: string[]
  dataType: "baile" | "cultural"
}

interface IndividualExcelParams {
  stats: CrossTabStats
  uniqueGenders: string[]
  title: string
  fileName: string
}

export function processCSVData(data: any[], type: "baile" | "cultural"): ProcessedData {
  console.log("Iniciando procesamiento de datos CSV")
  console.log("Número de filas:", data.length)
  console.log("Tipo de datos:", type)

  // Mostrar las columnas disponibles en la primera fila
  if (data.length > 0) {
    console.log("Columnas disponibles:", Object.keys(data[0]))
  }

  const participants: Participant[] = []
  const facultyStats: CrossTabStats = {}
  const programStats: CrossTabStats = {}
  const grupoCulturalStats: CrossTabStats = {}
  const statsByGrupoIndividual: { [grupo: string]: { byFaculty: CrossTabStats; byProgram: CrossTabStats } } = {}
  const genders = new Set<string>()

  data.forEach((row, index) => {
    // Saltar filas vacías
    if (!row || Object.keys(row).length === 0) {
      return
    }

    // Extraer el género usando la columna GENERO
    let genero = row["GENERO"] || row["Identidad de Género"] || "No especificado"
    if (typeof genero === "string") {
      genero = genero.trim()
    }
    if (!genero || genero === "") {
      genero = "No especificado"
    }
    genders.add(genero)

    // Extraer la facultad usando la columna FACULTAD O DEPENDENCIA A LA QUE PERTENECE
    let facultad =
      row["FACULTAD O DEPENDENCIA A LA QUE PERTENECE"] || row["DEPENDENCIA A LA QUE PERTENECE"] || "No especificada"
    if (typeof facultad === "string") {
      facultad = facultad.trim()
    }
    if (!facultad || facultad === "") {
      facultad = "No especificada"
    }

    // Determinar el programa académico buscando en las columnas de programas
    let programaAcademico = "No especificado"
    const facultyPrograms = [
      "PROGRAMAS DE LA FACULTAD DE ARTES INTEGRADAS",
      "PROGRAMAS DE LA FACULTAD DE CIENCIAS DE LA ADMINISTRACIÓN",
      "PROGRAMAS DE LA FACULTAD DE CIENCIAS NATURALES Y EXACTAS",
      "PROGRAMAS DE LA FACULTAD DE CIENCIAS SOCIALES Y ECONÓMICAS",
      "PROGRAMAS DE LA FACULTAD DE DERECHO Y CIENCIAS POLÍTICAS",
      "PROGRAMAS DE LA FACULTAD DE EDUCACIÓN Y PEDAGOGÍA",
      "PROGRAMAS DE LA FACULTAD DE HUMANIDADES",
      "PROGRAMAS DE LA FACULTAD DE INGENIERÍA",
      "PROGRAMAS DE LA FACULTAD DE PSICOLOGÍA",
      "PROGRAMAS DE LA FACULTAD DE SALUD",
    ]

    for (const field of facultyPrograms) {
      if (row[field] && typeof row[field] === "string" && row[field].trim() !== "") {
        programaAcademico = row[field].trim()
        break
      }
    }

    // Si no se encontró en las columnas específicas, buscar en otras posibles columnas
    if (programaAcademico === "No especificado") {
      const possibleProgramFields = ["Programas Académicos", "PROGRAMA ACADÉMICO", "Programa"]

      for (const field of possibleProgramFields) {
        if (row[field] && typeof row[field] === "string" && row[field].trim() !== "") {
          programaAcademico = row[field].trim()
          break
        }
      }
    }

    // Determinar el grupo cultural (solo para el tipo 'cultural')
    let grupoCultural = "No especificado"
    if (type === "cultural") {
      grupoCultural = row["GRUPOS CULTURALES"] || row["Grupo Cultural"] || row["GRUPO CULTURAL"] || "No especificado"
      if (typeof grupoCultural === "string") {
        grupoCultural = grupoCultural.trim()
      }
      if (!grupoCultural || grupoCultural === "") {
        grupoCultural = "No especificado"
      }
    }

    // Debug para las primeras filas
    if (index < 3) {
      console.log(`Fila ${index + 1}:`, {
        genero,
        facultad,
        programaAcademico,
        grupoCultural: type === "cultural" ? grupoCultural : "N/A",
      })
    }

    // Agregar participante a la lista
    const participant: Participant = {
      genero,
      facultad,
      programaAcademico,
    }

    if (type === "cultural") {
      participant.grupoCultural = grupoCultural
    }

    participants.push(participant)

    // Actualizar estadísticas por facultad
    if (!facultyStats[facultad]) {
      facultyStats[facultad] = { total: 0 }
    }
    if (!facultyStats[facultad][genero]) {
      facultyStats[facultad][genero] = 0
    }
    facultyStats[facultad][genero]++
    facultyStats[facultad].total++

    // Actualizar estadísticas por programa
    if (!programStats[programaAcademico]) {
      programStats[programaAcademico] = { total: 0 }
    }
    if (!programStats[programaAcademico][genero]) {
      programStats[programaAcademico][genero] = 0
    }
    programStats[programaAcademico][genero]++
    programStats[programaAcademico].total++

    // Actualizar estadísticas por grupo cultural (solo para el tipo 'cultural')
    if (type === "cultural") {
      if (!grupoCulturalStats[grupoCultural]) {
        grupoCulturalStats[grupoCultural] = { total: 0 }
      }
      if (!grupoCulturalStats[grupoCultural][genero]) {
        grupoCulturalStats[grupoCultural][genero] = 0
      }
      grupoCulturalStats[grupoCultural][genero]++
      grupoCulturalStats[grupoCultural].total++

      // Crear estadísticas individuales por grupo cultural
      if (!statsByGrupoIndividual[grupoCultural]) {
        statsByGrupoIndividual[grupoCultural] = {
          byFaculty: {},
          byProgram: {},
        }
      }

      // Estadísticas por facultad para este grupo
      if (!statsByGrupoIndividual[grupoCultural].byFaculty[facultad]) {
        statsByGrupoIndividual[grupoCultural].byFaculty[facultad] = { total: 0 }
      }
      if (!statsByGrupoIndividual[grupoCultural].byFaculty[facultad][genero]) {
        statsByGrupoIndividual[grupoCultural].byFaculty[facultad][genero] = 0
      }
      statsByGrupoIndividual[grupoCultural].byFaculty[facultad][genero]++
      statsByGrupoIndividual[grupoCultural].byFaculty[facultad].total++

      // Estadísticas por programa para este grupo
      if (!statsByGrupoIndividual[grupoCultural].byProgram[programaAcademico]) {
        statsByGrupoIndividual[grupoCultural].byProgram[programaAcademico] = { total: 0 }
      }
      if (!statsByGrupoIndividual[grupoCultural].byProgram[programaAcademico][genero]) {
        statsByGrupoIndividual[grupoCultural].byProgram[programaAcademico][genero] = 0
      }
      statsByGrupoIndividual[grupoCultural].byProgram[programaAcademico][genero]++
      statsByGrupoIndividual[grupoCultural].byProgram[programaAcademico].total++
    }
  })

  console.log("Procesamiento completado:")
  console.log("- Participantes:", participants.length)
  console.log("- Géneros únicos:", Array.from(genders))
  console.log("- Facultades:", Object.keys(facultyStats))
  console.log("- Programas:", Object.keys(programStats))
  if (type === "cultural") {
    console.log("- Grupos culturales:", Object.keys(grupoCulturalStats))
  }

  return {
    participants,
    statsByFaculty: facultyStats,
    statsByProgram: programStats,
    statsByGrupoCultural: grupoCulturalStats,
    statsByGrupoIndividual,
    uniqueGenders: Array.from(genders),
  }
}

// Función para generar Excel individual de una tabla
export async function generateIndividualExcel({
  stats,
  uniqueGenders,
  title,
  fileName,
}: IndividualExcelParams): Promise<void> {
  const workbook = new ExcelJS.Workbook()
  workbook.creator = "Universidad del Valle"
  workbook.created = new Date()

  const worksheet = workbook.addWorksheet(title)

  // Configurar encabezados
  worksheet.addRow([title])
  const cellA1 = worksheet.getCell("A1")
  cellA1.alignment = { horizontal: "center", vertical: "middle" }
  worksheet.mergeCells("A1:A2")

  // Añadir encabezado "Género"
  const genderHeaderCell = worksheet.getCell(`B1`)
  genderHeaderCell.value = "Género"
  genderHeaderCell.alignment = { horizontal: "center" }
  worksheet.mergeCells(`B1:${String.fromCharCode(65 + uniqueGenders.length)}1`)

  // Añadir columna "Total"
  const totalHeaderCell = worksheet.getCell(`${String.fromCharCode(66 + uniqueGenders.length)}1`)
  totalHeaderCell.value = "Total"
  totalHeaderCell.alignment = { horizontal: "center", vertical: "middle" }
  worksheet.mergeCells(
    `${String.fromCharCode(66 + uniqueGenders.length)}1:${String.fromCharCode(66 + uniqueGenders.length)}2`,
  )

  // Añadir encabezados de género en la segunda fila
  const genderRow = worksheet.getRow(2)
  uniqueGenders.forEach((gender, index) => {
    const cell = genderRow.getCell(index + 2)
    cell.value = gender
    cell.alignment = { horizontal: "center" }
  })

  // Añadir datos
  let rowIndex = 3
  Object.entries(stats).forEach(([key, stat]) => {
    const row = worksheet.addRow([key])

    uniqueGenders.forEach((gender, index) => {
      const cell = row.getCell(index + 2)
      cell.value = stat[gender] || 0
      cell.alignment = { horizontal: "center" }
    })

    const totalCell = row.getCell(uniqueGenders.length + 2)
    totalCell.value = stat.total || 0
    totalCell.alignment = { horizontal: "center" }
    totalCell.font = { bold: true }

    if (rowIndex % 2 === 1) {
      row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } }
    }

    rowIndex++
  })

  // Añadir fila de totales
  const totalRow = worksheet.addRow(["TOTAL"])
  totalRow.font = { bold: true }

  uniqueGenders.forEach((gender, index) => {
    const total = Object.values(stats).reduce((sum, stat) => sum + (stat[gender] || 0), 0)
    const cell = totalRow.getCell(index + 2)
    cell.value = total
    cell.alignment = { horizontal: "center" }
  })

  const grandTotal = Object.values(stats).reduce((sum, stat) => sum + (stat.total || 0), 0)
  const grandTotalCell = totalRow.getCell(uniqueGenders.length + 2)
  grandTotalCell.value = grandTotal
  grandTotalCell.alignment = { horizontal: "center" }

  // Aplicar estilos
  const row1 = worksheet.getRow(1)
  row1.font = { bold: true, color: { argb: "FFFFFFFF" } }
  row1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F81BD" } }
  row1.alignment = { horizontal: "center" }

  const row2 = worksheet.getRow(2)
  row2.font = { bold: true, color: { argb: "FFFFFFFF" } }
  row2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F81BD" } }
  row2.alignment = { horizontal: "center" }

  totalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } }

  // Ajustar ancho de columnas
  if (worksheet.columns) {
    worksheet.columns.forEach((column) => {
      if (!column) return
      let maxLength = 0
      column.eachCell({ includeEmpty: true }, (cell) => {
        if (!cell) return
        const cellValue = cell.value
        const columnLength = cellValue ? cellValue.toString().length : 10
        if (columnLength > maxLength) {
          maxLength = columnLength
        }
      })
      column.width = Math.min(maxLength + 2, 50)
    })
  }

  // Generar y descargar archivo
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = fileName
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

export async function generateExcel({
  statsByFaculty,
  statsByProgram,
  statsByGrupoCultural,
  statsByGrupoIndividual,
  uniqueGenders,
  dataType,
}: ExcelGenerationParams): Promise<void> {
  // Crear un nuevo libro de Excel
  const workbook = new ExcelJS.Workbook()
  workbook.creator = "Universidad del Valle"
  workbook.created = new Date()

  // Función para aplicar estilos a los encabezados
  const applyHeaderStyles = (worksheet: ExcelJS.Worksheet) => {
    // Estilo para encabezados
    const row1 = worksheet.getRow(1)
    row1.font = { bold: true, color: { argb: "FFFFFFFF" } }
    row1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F81BD" } }
    row1.alignment = { horizontal: "center" }

    // Si hay una segunda fila de encabezados (para géneros)
    const row2 = worksheet.getRow(2)
    const cell = row2.getCell(1)
    if (cell && cell.value) {
      row2.font = { bold: true, color: { argb: "FFFFFFFF" } }
      row2.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4F81BD" } }
      row2.alignment = { horizontal: "center" }
    }
  }

  // Función para aplicar estilos a la fila de totales
  const applyTotalRowStyles = (worksheet: ExcelJS.Worksheet, rowIndex: number) => {
    const totalRow = worksheet.getRow(rowIndex)
    totalRow.font = { bold: true }
    totalRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9D9D9" } }
  }

  // Función para ajustar el ancho de las columnas
  const adjustColumnWidths = (worksheet: ExcelJS.Worksheet) => {
    if (!worksheet.columns) return

    worksheet.columns.forEach((column) => {
      if (!column) return

      let maxLength = 0
      column.eachCell({ includeEmpty: true }, (cell) => {
        if (!cell) return
        const cellValue = cell.value
        const columnLength = cellValue ? cellValue.toString().length : 10
        if (columnLength > maxLength) {
          maxLength = columnLength
        }
      })
      column.width = Math.min(maxLength + 2, 50) // Limitar a 50 para evitar columnas demasiado anchas
    })
  }

  // Crear hoja para estadísticas por facultad
  const facultySheet = workbook.addWorksheet("Estadísticas por Facultad")

  // Configurar encabezados con merge para "Género"
  facultySheet.addRow(["Facultad"])
  const cellA1 = facultySheet.getCell("A1")
  cellA1.alignment = { horizontal: "center", vertical: "middle" }
  facultySheet.mergeCells("A1:A2")

  // Añadir encabezado "Género" que abarca todas las columnas de género
  const genderHeaderCell = facultySheet.getCell(`B1`)
  genderHeaderCell.value = "Género"
  genderHeaderCell.alignment = { horizontal: "center" }
  facultySheet.mergeCells(`B1:${String.fromCharCode(65 + uniqueGenders.length)}1`) // B1:C1, B1:D1, etc.

  // Añadir columna "Total"
  const totalHeaderCell = facultySheet.getCell(`${String.fromCharCode(66 + uniqueGenders.length)}1`)
  totalHeaderCell.value = "Total"
  totalHeaderCell.alignment = { horizontal: "center", vertical: "middle" }
  facultySheet.mergeCells(
    `${String.fromCharCode(66 + uniqueGenders.length)}1:${String.fromCharCode(66 + uniqueGenders.length)}2`,
  )

  // Añadir encabezados de género en la segunda fila
  const genderRow = facultySheet.getRow(2)
  uniqueGenders.forEach((gender, index) => {
    const cell = genderRow.getCell(index + 2)
    cell.value = gender
    cell.alignment = { horizontal: "center" }
  })

  // Añadir datos de facultad
  let rowIndex = 3
  Object.entries(statsByFaculty).forEach(([faculty, stats]) => {
    const row = facultySheet.addRow([faculty])

    uniqueGenders.forEach((gender, index) => {
      const cell = row.getCell(index + 2)
      cell.value = stats[gender] || 0
      cell.alignment = { horizontal: "center" }
    })

    const totalCell = row.getCell(uniqueGenders.length + 2)
    totalCell.value = stats.total || 0
    totalCell.alignment = { horizontal: "center" }
    totalCell.font = { bold: true }

    // Alternar colores de fila
    if (rowIndex % 2 === 1) {
      row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } }
    }

    rowIndex++
  })

  // Añadir fila de totales
  const facultyTotalRow = facultySheet.addRow(["TOTAL"])
  facultyTotalRow.font = { bold: true }

  uniqueGenders.forEach((gender, index) => {
    const total = Object.values(statsByFaculty).reduce((sum, stats) => sum + (stats[gender] || 0), 0)
    const cell = facultyTotalRow.getCell(index + 2)
    cell.value = total
    cell.alignment = { horizontal: "center" }
  })

  const facultyGrandTotal = Object.values(statsByFaculty).reduce((sum, stats) => sum + (stats.total || 0), 0)
  const grandTotalCell = facultyTotalRow.getCell(uniqueGenders.length + 2)
  grandTotalCell.value = facultyGrandTotal
  grandTotalCell.alignment = { horizontal: "center" }

  // Aplicar estilos a la fila de totales
  applyTotalRowStyles(facultySheet, rowIndex)

  // Aplicar estilos a los encabezados
  applyHeaderStyles(facultySheet)

  // Ajustar ancho de columnas
  adjustColumnWidths(facultySheet)

  // Crear hoja para estadísticas por programa
  const programSheet = workbook.addWorksheet("Estadísticas por Programa")

  // Configurar encabezados con merge para "Género"
  programSheet.addRow(["Programa Académico"])
  const programCellA1 = programSheet.getCell("A1")
  programCellA1.alignment = { horizontal: "center", vertical: "middle" }
  programSheet.mergeCells("A1:A2")

  // Añadir encabezado "Género" que abarca todas las columnas de género
  const programGenderHeaderCell = programSheet.getCell(`B1`)
  programGenderHeaderCell.value = "Género"
  programGenderHeaderCell.alignment = { horizontal: "center" }
  programSheet.mergeCells(`B1:${String.fromCharCode(65 + uniqueGenders.length)}1`) // B1:C1, B1:D1, etc.

  // Añadir columna "Total"
  const programTotalHeaderCell = programSheet.getCell(`${String.fromCharCode(66 + uniqueGenders.length)}1`)
  programTotalHeaderCell.value = "Total"
  programTotalHeaderCell.alignment = { horizontal: "center", vertical: "middle" }
  programSheet.mergeCells(
    `${String.fromCharCode(66 + uniqueGenders.length)}1:${String.fromCharCode(66 + uniqueGenders.length)}2`,
  )

  // Añadir encabezados de género en la segunda fila
  const programGenderRow = programSheet.getRow(2)
  uniqueGenders.forEach((gender, index) => {
    const cell = programGenderRow.getCell(index + 2)
    cell.value = gender
    cell.alignment = { horizontal: "center" }
  })

  // Añadir datos de programa
  let programRowIndex = 3
  Object.entries(statsByProgram).forEach(([program, stats]) => {
    const row = programSheet.addRow([program])

    uniqueGenders.forEach((gender, index) => {
      const cell = row.getCell(index + 2)
      cell.value = stats[gender] || 0
      cell.alignment = { horizontal: "center" }
    })

    const totalCell = row.getCell(uniqueGenders.length + 2)
    totalCell.value = stats.total || 0
    totalCell.alignment = { horizontal: "center" }
    totalCell.font = { bold: true }

    // Alternar colores de fila
    if (programRowIndex % 2 === 1) {
      row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } }
    }

    programRowIndex++
  })

  // Añadir fila de totales
  const programTotalRow = programSheet.addRow(["TOTAL"])
  programTotalRow.font = { bold: true }

  uniqueGenders.forEach((gender, index) => {
    const total = Object.values(statsByProgram).reduce((sum, stats) => sum + (stats[gender] || 0), 0)
    const cell = programTotalRow.getCell(index + 2)
    cell.value = total
    cell.alignment = { horizontal: "center" }
  })

  const programGrandTotal = Object.values(statsByProgram).reduce((sum, stats) => sum + (stats.total || 0), 0)
  const programGrandTotalCell = programTotalRow.getCell(uniqueGenders.length + 2)
  programGrandTotalCell.value = programGrandTotal
  programGrandTotalCell.alignment = { horizontal: "center" }

  // Aplicar estilos a la fila de totales
  applyTotalRowStyles(programSheet, programRowIndex)

  // Aplicar estilos a los encabezados
  applyHeaderStyles(programSheet)

  // Ajustar ancho de columnas
  adjustColumnWidths(programSheet)

  // Crear hoja para estadísticas por grupo cultural (si es del tipo 'cultural')
  if (dataType === "cultural" && Object.keys(statsByGrupoCultural).length > 0) {
    const culturalSheet = workbook.addWorksheet("Estadísticas por Grupo Cultural")

    // Configurar encabezados con merge para "Género"
    culturalSheet.addRow(["Grupo Cultural"])
    const culturalCellA1 = culturalSheet.getCell("A1")
    culturalCellA1.alignment = { horizontal: "center", vertical: "middle" }
    culturalSheet.mergeCells("A1:A2")

    // Añadir encabezado "Género" que abarca todas las columnas de género
    const culturalGenderHeaderCell = culturalSheet.getCell(`B1`)
    culturalGenderHeaderCell.value = "Género"
    culturalGenderHeaderCell.alignment = { horizontal: "center" }
    culturalSheet.mergeCells(`B1:${String.fromCharCode(65 + uniqueGenders.length)}1`) // B1:C1, B1:D1, etc.

    // Añadir columna "Total"
    const culturalTotalHeaderCell = culturalSheet.getCell(`${String.fromCharCode(66 + uniqueGenders.length)}1`)
    culturalTotalHeaderCell.value = "Total"
    culturalTotalHeaderCell.alignment = { horizontal: "center", vertical: "middle" }
    culturalSheet.mergeCells(
      `${String.fromCharCode(66 + uniqueGenders.length)}1:${String.fromCharCode(66 + uniqueGenders.length)}2`,
    )

    // Añadir encabezados de género en la segunda fila
    const culturalGenderRow = culturalSheet.getRow(2)
    uniqueGenders.forEach((gender, index) => {
      const cell = culturalGenderRow.getCell(index + 2)
      cell.value = gender
      cell.alignment = { horizontal: "center" }
    })

    // Añadir datos de grupo cultural
    let culturalRowIndex = 3
    Object.entries(statsByGrupoCultural).forEach(([group, stats]) => {
      const row = culturalSheet.addRow([group])

      uniqueGenders.forEach((gender, index) => {
        const cell = row.getCell(index + 2)
        cell.value = stats[gender] || 0
        cell.alignment = { horizontal: "center" }
      })

      const totalCell = row.getCell(uniqueGenders.length + 2)
      totalCell.value = stats.total || 0
      totalCell.alignment = { horizontal: "center" }
      totalCell.font = { bold: true }

      // Alternar colores de fila
      if (culturalRowIndex % 2 === 1) {
        row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF2F2F2" } }
      }

      culturalRowIndex++
    })

    // Añadir fila de totales
    const culturalTotalRow = culturalSheet.addRow(["TOTAL"])
    culturalTotalRow.font = { bold: true }

    uniqueGenders.forEach((gender, index) => {
      const total = Object.values(statsByGrupoCultural).reduce((sum, stats) => sum + (stats[gender] || 0), 0)
      const cell = culturalTotalRow.getCell(index + 2)
      cell.value = total
      cell.alignment = { horizontal: "center" }
    })

    const culturalGrandTotal = Object.values(statsByGrupoCultural).reduce((sum, stats) => sum + (stats.total || 0), 0)
    const grandTotalCell = culturalTotalRow.getCell(uniqueGenders.length + 2)
    grandTotalCell.value = culturalGrandTotal
    grandTotalCell.alignment = { horizontal: "center" }

    // Aplicar estilos a la fila de totales
    applyTotalRowStyles(culturalSheet, culturalRowIndex)

    // Aplicar estilos a los encabezados
    applyHeaderStyles(culturalSheet)

    // Ajustar ancho de columnas
    adjustColumnWidths(culturalSheet)
  }

  // Generar archivo Excel y descargarlo
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
  const url = URL.createObjectURL(blob)
  const a = document.createElement("a")
  a.href = url
  a.download = `Estadisticas_${dataType === "baile" ? "Baile_Recreativo" : "Grupos_Culturales"}.xlsx`
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
