interface CrossTabStats {
  [key: string]: {
    [gender: string]: number
    total: number
  }
}

interface StatsTableProps {
  title: string
  stats: CrossTabStats
  genders: string[]
}

export function StatsTable({ title, stats, genders }: StatsTableProps) {
  // Calcular totales por género
  const genderTotals: { [key: string]: number } = {}
  let grandTotal = 0

  genders.forEach((gender) => {
    const total = Object.values(stats).reduce((sum, stat) => sum + (stat[gender] || 0), 0)
    genderTotals[gender] = total
    grandTotal += total
  })

  return (
    <div className="overflow-x-auto">
      <table className="w-full border-collapse">
        <thead>
          <tr className="bg-muted">
            <th className="border p-2 text-left" rowSpan={2}>
              {title}
            </th>
            <th className="border p-2 text-center" colSpan={genders.length}>
              Género
            </th>
            <th className="border p-2 text-center" rowSpan={2}>
              Total
            </th>
          </tr>
          <tr className="bg-muted">
            {genders.map((gender) => (
              <th key={gender} className="border p-2 text-center">
                {gender}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {Object.entries(stats).map(([key, stat], index) => (
            <tr key={key} className={index % 2 === 0 ? "bg-white" : "bg-muted/30"}>
              <td className="border p-2">{key}</td>
              {genders.map((gender) => (
                <td key={`${key}-${gender}`} className="border p-2 text-center">
                  {stat[gender] || 0}
                </td>
              ))}
              <td className="border p-2 text-center font-medium">{stat.total || 0}</td>
            </tr>
          ))}
          <tr className="bg-muted font-medium">
            <td className="border p-2">TOTAL</td>
            {genders.map((gender) => (
              <td key={`total-${gender}`} className="border p-2 text-center">
                {genderTotals[gender]}
              </td>
            ))}
            <td className="border p-2 text-center">{grandTotal}</td>
          </tr>
        </tbody>
      </table>
    </div>
  )
}
