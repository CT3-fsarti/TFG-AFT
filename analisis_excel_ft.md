# Analisis del Excel Diseño Red FT v5a.xlsm

## Resumen ejecutivo

El libro implementa un modelo dirigido y ponderado de una red de financiacion del terrorismo en 4 capas logicas:

1. Definicion de taxonomia y nodos.
2. Definicion de canales y conversion de etiquetas cualitativas en pesos numericos.
3. Construccion de enlaces agregados y matrices de red.
4. Calculo de distancias y metricas de nodos y de red.

La estructura actual contiene:

- 16 hojas.
- 31 nodos.
- 56 canales.
- 217 filas en la tabla `tblEnlaces` con combinaciones nodo inicial - nodo final - canal.
- 58 enlaces unicos agregados (`E001` ... `E058`).
- 50 enlaces con mas de un canal posible.
- Un maximo de 13 canales asociados a un mismo enlace.

La hoja de distancias minimas no esta calculada con formulas normales de Excel en las celdas interiores: guarda valores ya materializados. Para la version Python conviene recalcularla directamente con un algoritmo de caminos minimos, no intentar reproducir ese almacenamiento intermedio.

## Inventario de hojas y tablas

### 1. Portada

- Tabla: `tblIndiceHojas` (`B3:D18`)
- Funcion: indice y navegacion documental.

### 2. Tipos de Nodo

- Tabla: `tblTiposDeNodo` (`B3:E8`)
- Columnas: `Tipo`, `Descripcion`, `Capa`, `Notas`
- Tipos definidos:
  - `O`: Origen
  - `C`: Colector
  - `I`: Intermediario
  - `G`: Gestor
  - `D`: Destino

### 3. Nodos

- Tabla: `tblNodos` (`B3:E34`)
- Columnas: `NodoID`, `Nombre`, `Tipo`, `Activo`
- Conteo por tipo:
  - `O`: 8
  - `C`: 6
  - `I`: 12
  - `G`: 4
  - `D`: 1

### 4. Prop.&Pesos Canales

Tablas de parametrizacion:

- `tblPesosCoste` (`B3:C8`)
  - `Valor` -> `Peso`
  - `Bajo=5`, `Bajo-Medio=4`, `Medio=3`, `Medio-Alto=2`, `Alto=1`

- `tblPesosValorOperativo` (`E3:J9`)
  - Columnas: `Valor`, `Frecuencia de Uso`, `Opacidad`, `Trazabilidad`, `Velocidad`, `Escalabilidad`
  - Escala:
    - `Bajo = (1, 1, 5, 1, 1)`
    - `Bajo-Medio = (2, 2, 4, 2, 2)`
    - `Medio = (3, 3, 3, 3, 3)`
    - `Medio-Alto = (4, 4, 2, 4, 4)`
    - `Alto = (5, 5, 1, 5, 5)`

- `tblRangoEnlace` (`B10:C12`)
  - Rango posible de coste agregado: minimo 1, maximo 5.

- `tblRangoEnlace1ValorOpertivo` (`E10:J12`)
  - Rango posible del valor operativo agregado: minimo 4, maximo 20.

Observacion clave: la trazabilidad puntua al reves que opacidad, velocidad y escalabilidad. Menor trazabilidad es mejor para la red FT y por eso recibe mayor puntuacion.

### 5. Canales

- `tblPropiedadesCanales` (`B3:C8`)
  - Tabla descriptiva de variables.

- `tblCanales` (`E3:P59`)
  - Columnas manuales:
    - `Canal`
    - `Frecuencia de Uso`
    - `Coste`
    - `Opacidad`
    - `Trazabilidad`
    - `Velocidad`
    - `Escalabilidad`
    - textos explicativos de cada dimension.

Esta hoja funciona como catalogo cualitativo de mecanismos de transferencia. Sus etiquetas se traducen a pesos numericos al construir `tblEnlaces`.

### 6. Enlaces

- Tabla: `tblEnlaces` (`B3:P436`)
- Es la hoja central del modelo.
- Cada fila representa un canal posible entre un nodo inicial y un nodo final.

Columnas manuales:

- `Nodo Inicial`
- `Nodo Final`
- `Canal`
- `Activo`

Columnas derivadas:

- `Enlace`
  - Formula derramada que asigna un identificador `E###` a cada pareja unica `(Nodo Inicial, Nodo Final)` segun su primer orden de aparicion.
  - Formula base:

```excel
=IF(OR([Nodo Inicial]="",[Nodo Final]=""),"","E"&TEXT(
  XMATCH([Nodo Inicial]&"|"&[Nodo Final],
    UNIQUE(FILTER(tblEnlaces[Nodo Inicial]&"|"&tblEnlaces[Nodo Final],
      (tblEnlaces[Nodo Inicial]<>"")*(tblEnlaces[Nodo Final]<>""))),0),
"000"))
```

- `Detalle Nodos Iniciales`
  - `VLOOKUP` contra `tblNodos` para obtener el nombre del nodo inicial.

- `Detalle Nodos Finales`
  - `VLOOKUP` contra `tblNodos` para obtener el nombre del nodo final.

- `Frecuencia de Uso`
  - Convierte la etiqueta textual del canal a peso numerico usando `tblPesosValorOperativo[Frecuencia de Uso]`.

- `Frecuencia de Uso Relativa`
  - Normaliza la frecuencia dentro del mismo enlace solo con filas activas.

$$
f^{rel}_{i} = \frac{f_i}{\sum_{j \in enlace, activo} f_j}
$$

- `COSTE`
  - Convierte la etiqueta textual de coste del canal a peso numerico usando `tblPesosCoste`.

- `Opacidad`, `Trazabilidad`, `Velocidad`, `Escalabilidad`
  - Convierten sus etiquetas textuales en pesos numericos usando `tblPesosValorOperativo`.

- `VALOR OPERATIVO`

$$
VO = Opacidad + Trazabilidad + Velocidad + Escalabilidad
$$

### 7. Red

- Tabla: `tblRed` (`D3:J61`)
- Consolida `tblEnlaces` a nivel de enlace unico.

Helper fuera de tabla:

- `B4`: `SORT(UNIQUE(FILTER(tblEnlaces[Enlace], tblEnlaces[Enlace]<>"")))`

Columnas derivadas:

- `Enlace`
  - Toma la lista dinamica de IDs unicos desde el helper derramado.

- `Nodo Inicial`, `Nodo Final`
  - `VLOOKUP` sobre `tblEnlaces`.

- `Activo`
  - Vale 0 si:
    - el nodo inicial esta inactivo, o
    - el nodo final esta inactivo, o
    - todas las filas de `tblEnlaces` asociadas a ese `Enlace` estan inactivas.
  - En otro caso vale 1.

- `COSTE`
  - Media ponderada por frecuencia relativa de los costes de los canales activos del enlace.

$$
Coste_{enlace} = \sum_i Coste_i \cdot f^{rel}_i
$$

- `VALOR OPERATIVO`
  - Media ponderada por frecuencia relativa del valor operativo de los canales activos del enlace.

$$
VO_{enlace} = \sum_i VO_i \cdot f^{rel}_i
$$

- `TRADE OFF VALOR OPERATIVO/COSTE`

$$
TradeOff = \frac{VO_{enlace}}{Coste_{enlace}}
$$

Ejemplos reales:

- `E001`: `O1 -> C4`, coste `5`, valor operativo `9`, trade-off `1.8`
- `E002`: `O1 -> I5`, coste `4.25`, valor operativo `10.5`, trade-off `2.470588...`

### 8. Grafo Visual

- Sin formulas ni tablas utiles para el calculo estructural.
- Parece reservada para visualizacion.

### 9. MatrizDeAdyacencia

- Tabla: `tblMatrizDeAdyacencia` (`B3:AH34`)
- Filas y columnas indexadas por los 31 nodos.
- Valor interior:
  - `1` si existe un enlace activo entre origen y destino.
  - `0` en caso contrario.

Formula tipo:

```excel
=IF(COUNTIFS(tblRed[Nodo Inicial], fila_nodo,
             tblRed[Nodo Final], columna_nodo,
             tblRed[Activo],1)>0,1,0)
```

### 10. MatrizPonderadaCostes

- Tabla: `tblMatrizPonderadaCostes` (`B3:AH34`)
- Copia la topologia de adyacencia y coloca el coste agregado del enlace activo.

Formula tipo:

```excel
=IF(MatrizDeAdyacencia!celda=1,
    SUMIFS(tblRed[COSTE], tblRed[Nodo Inicial], fila_nodo,
                         tblRed[Nodo Final], columna_nodo,
                         tblRed[Activo], 1),
    0)
```

### 11. MatrizPonderadaValorOperativo

- Tabla: `tblMatrizPonderadaValorOperativo` (`B3:AH34`)
- Igual que la anterior pero usando `VALOR OPERATIVO`.

### 12. MatrizTradeOff Valor Op-Coste

- Tabla: `tblMatrizTradeOff` (`B3:AH34`)
- Igual que la anterior pero usando `TRADE OFF VALOR OPERATIVO/COSTE`.
- Esta es la matriz ponderada principal para las metricas de centralidad.

### 13. MatrizDeDistanciasDirectas

- Tabla: `tblMatrizDistanciasDirectas` (`B3:AH34`)
- Convierte el trade-off en una distancia o friccion.

Regla:

- diagonal -> `0`
- sin enlace -> `999999`
- con enlace -> inverso del trade-off

$$
d_{ij}^{directa} =
\begin{cases}
0 & \text{si } i=j \\
999999 & \text{si no hay enlace} \\
\frac{1}{TradeOff_{ij}} & \text{si hay enlace}
\end{cases}
$$

### 14. Métricas de Nodos

- Tabla: `tblMetricasNodos` (`B3:H35`)
- Columnas:
  - `Nombre Nodo`
  - `Nodo`
  - `Grado Ponderado de Entrada`
  - `Grado Ponderado de Salida`
  - `Grado Ponderado Total`
  - `Grado Normalizado`
  - `Cercanía Armónica`

Formulas equivalentes:

- Grado ponderado de entrada:

$$
In_i = \sum_j TradeOff_{j,i}
$$

- Grado ponderado de salida:

$$
Out_i = \sum_j TradeOff_{i,j}
$$

- Grado ponderado total:

$$
Deg_i = In_i + Out_i
$$

- Grado normalizado:

$$
DegNorm_i = \frac{Deg_i}{\sum_k Deg_k}
$$

- Cercania armonica:

$$
HC_i = \sum_{j \neq i, d_{ij}<999999} \frac{1}{d_{ij}}
$$

### 15. MatrizDeDistanciasMínimas

- Tabla: `tblMatrizDistanciasMínimas` (`B3:AH34`)
- Los encabezados y etiquetas laterales usan `VLOOKUP`, pero las celdas interiores ya contienen valores numericos materializados.
- No aparecen formulas normales en el cuerpo de la matriz.
- El libro contiene `vbaProject.bin` y un modulo VBA real llamado `Módulo1.bas` con la macro `CalcularMatrizDistanciasMinimas()`.
- Esa macro:
  - toma como origen la tabla `tblMatrizDistanciasDirectas`
  - toma como destino la tabla `tblMatrizDistanciasMínimas`
  - extrae solo el bloque numerico de ambas tablas, excluyendo la primera fila y las dos primeras columnas
  - usa `INF = 999999`
  - ejecuta Floyd-Warshall clasico triple bucle
  - escribe el resultado en la tabla destino

Equivalente VBA confirmado:

```vb
For k = 1 To nFilas
  For i = 1 To nFilas
    For j = 1 To nCols
      If dist(i, k) < INF And dist(k, j) < INF Then
        candidato = dist(i, k) + dist(k, j)
        If candidato < dist(i, j) Then
          dist(i, j) = candidato
        End If
      End If
    Next j
  Next i
Next k
```

Implicacion para Python:

- No intentes leer esta hoja como fuente de verdad.
- Recalcula caminos minimos desde `MatrizDeDistanciasDirectas` o directamente desde los enlaces de `tblRed`.
- Si quieres reproducir el comportamiento exacto del Excel, usa Floyd-Warshall sobre una matriz densa de `31 x 31` con `np.inf` como sustituto interno de `999999`.

### 16. Métricas de Red

- Tabla: `tblMetricasNodos20` (`B3:E8`)
- Metricas finales observadas:

1. Distancia media al destino final `D1`
2. Distancia media total de la red
3. Cercania armonica media de la red
4. Eficiencia global de la red
5. Centralizacion de la red

Formulas equivalentes:

- Distancia media al destino final:

$$
\overline{d}_{\to D1} = mean(d_{i,D1} \; | \; 0 < d_{i,D1} < 999999)
$$

- Distancia media total:

$$
\overline{d} = mean(d_{ij} \; | \; 0 < d_{ij} < 999999)
$$

- Cercania armonica media:

$$
\overline{HC} = mean(HC_i)
$$

- Eficiencia global:

$$
GE = \frac{\sum_{i \neq j, d_{ij}<999999} 1/d_{ij}}{n(n-1)}
$$

- Centralizacion:

$$
Centralizacion = \frac{\sum_i (\max(Deg) - Deg_i)}{(n-1)\max(Deg)}
$$

Valores actualmente guardados en el libro:

- Distancia media al destino final `D1`: `1.0149661598669444`
- Distancia media total de la red: `0.7893366899612394`
- Cercania armonica media de la red: `6.821286841605575`
- Eficiencia global de la red: `0.22737622805351917`
- Centralizacion de la red: `0.7552351897610544`

## Flujo de definicion de la red FT

El proceso de modelado en el Excel sigue esta secuencia conceptual:

1. Definir la taxonomia de actores.
   - Los nodos se clasifican en origen, colector, intermediario, gestor y destino.

2. Instanciar los nodos concretos.
   - Cada nodo tiene un identificador, un nombre y un estado `Activo`.
   - La activacion permite simular intervenciones o neutralizaciones.

3. Definir un catalogo de canales financieros.
   - Cada canal se describe cualitativamente por coste, frecuencia de uso, opacidad, trazabilidad, velocidad y escalabilidad.

4. Traducir etiquetas cualitativas a pesos numericos.
   - Esta capa es la metodologia de scoring del modelo.

5. Construir filas de enlace-canal.
   - Para una misma pareja de nodos puede haber varios canales alternativos.
   - Cada fila recibe puntuaciones numericas y un identificador comun de enlace.

6. Agregar por enlace unico.
   - El enlace agregado resume varios canales posibles mediante medias ponderadas por frecuencia relativa.

7. Convertir la red agregada en matrices.
   - Adyacencia.
   - Coste ponderado.
   - Valor operativo ponderado.
   - Trade-off valor operativo / coste.
   - Distancias directas.

8. Calcular accesibilidad estructural.
   - A partir de las distancias directas se calculan caminos minimos entre todos los pares de nodos.

9. Extraer metricas.
   - Importancia relativa de nodos: grados ponderados y cercania armonica.
   - Desempeño global: distancias medias, eficiencia y centralizacion.

## Flujo de calculo recomendado en Python

La replica en Python deberia separar claramente entrada, transformacion y metricas.

### Librerias recomendadas

- `pandas` para cargar y transformar tablas.
- `networkx` para representar el grafo y calcular caminos minimos.
- `numpy` para operar con matrices y mascaras.
- `openpyxl` solo si quieres seguir leyendo o escribiendo Excel.

### Pipeline recomendado

1. Cargar tablas base:
   - `tblTiposDeNodo`
   - `tblNodos`
   - `tblPesosCoste`
   - `tblPesosValorOperativo`
   - `tblCanales`
   - entradas manuales de `tblEnlaces`

2. Normalizar texto:
   - mayusculas/minusculas
   - tildes opcionales
   - guiones tipo `Bajo-Medio` frente a `bajo-medio`

3. Construir `enlaces_detalle`:
   - mapear nombres de nodos
   - mapear pesos de coste y valor operativo
   - calcular `frecuencia_uso_relativa`
   - calcular `valor_operativo`
   - generar `enlace_id`

4. Construir `red_agregada` agrupando por `enlace_id`.

5. Construir matriz de trade-off y matriz de distancias directas.

6. Calcular caminos minimos all-pairs.
   - Con `networkx.all_pairs_dijkstra_path_length` usando `distancia_directa` como peso.
   - O con `floyd_warshall_numpy` si prefieres una salida matricial directa.

7. Calcular metricas de nodo y de red.

8. Exportar resultados a DataFrames o a Excel si necesitas mantener el formato de salida.

## Observaciones tecnicas relevantes

- El libro usa muchas formulas modernas de Excel (`XLOOKUP`, `XMATCH`, `UNIQUE`, `FILTER`, `LET`, `ANCHORARRAY`).
- En Python eso no es un problema: esas formulas se traducen mejor a operaciones de `merge`, `groupby`, `map` y calculos matriciales.
- La constante `999999` se usa como centinela de no alcanzabilidad. En Python es preferible usar `np.inf` internamente y solo convertir a `999999` al exportar si quieres reproducir exactamente el Excel.
- La tabla `Red` depende de una lista derramada auxiliar. En Python esto se sustituye naturalmente por `drop_duplicates` o `groupby`.
- La app actual de Streamlit no esta alineada del todo con el esquema de este workbook: sigue buscando nombres como `tblPesos`, `tblMatrizDistancias` y `tblGradoPonderado`, que no existen en `v5a`.

## Conclusión

El Excel no es un conjunto de hojas independientes, sino una cadena de transformacion bien definida:

`Nodos + Canales + Pesos -> Enlaces detallados -> Red agregada -> Matrices -> Distancias -> Métricas`

La parte mas importante para portar a Python no es reproducir cada formula celda a celda, sino preservar estas reglas:

1. Los canales se puntuan con una taxonomia cualitativa convertida a pesos numericos.
2. Varios canales pueden coexistir dentro de un mismo enlace entre nodos.
3. El enlace agregado se resume por medias ponderadas con frecuencia relativa.
4. El trade-off se transforma en distancia via inverso.
5. Las metricas se calculan sobre la matriz de trade-off y sobre los caminos minimos.

Si el siguiente paso es implementar el portado, la replica mas fiel y mantenible es un pipeline `pandas + networkx`, no una emulacion literal de formulas de Excel.

## Utilidad implementada

En el repositorio ya existe una capa intermedia funcional en `ft_excel_bridge.py`.

Que hace:

- carga y valida las tablas maestras del workbook
- reconstruye en Python `tblEnlaces`, `tblRed`, todas las matrices y todas las metricas
- replica tambien la parte VBA de Floyd-Warshall para `tblMatrizDistanciasMinimas`
- compara los resultados generados contra las tablas calculadas del Excel

Ejecucion recomendada:

```bash
.venv\Scripts\python.exe ft_excel_bridge.py --workbook "Diseño Red FT v5a.xlsm" --compare --json
```

Estado de validacion actual:

- las 9 salidas calculadas comparadas contra el workbook cuadran sin diferencias materiales
- los maximos errores absolutos residuales observados son de precision flotante de orden `1e-15`