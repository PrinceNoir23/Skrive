CompareMaps(mapA, mapB) {
    if Type(mapA) != "Map" || Type(mapB) != "Map" {
        MsgBox "Error: Uno de los valores no es un Map()"
        return
    }

    result := ""

    ; Verificar claves y valores en mapA que no coincidan en mapB
    for key in mapA {
        valueA := mapA[key]
        if !mapB.Has(key) {
            result .= "Clave faltante en Map2: " key "`n"
        } else {
            valueB := mapB[key]
            if (valueA != valueB) {
                result .= "Diferencia en " key ": " valueA " ≠ " valueB "`n"
            }
        }
    }

    ; Verificar claves que están en mapB pero no en mapA
    for key in mapB {
        if !mapA.Has(key)
            result .= "Clave faltante en Map1: " key "`n"
    }

    if result = ""
        result := "No hay diferencias"

    MsgBox result
}
