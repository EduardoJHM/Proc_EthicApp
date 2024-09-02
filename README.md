# Proc_EthicApp
Código en Rstudio con el cual se procesa la información obtenida en la realización de los casos en los cursos de plan común de ingeniería y ciencias de la FCFM 


¿Cómo funciona el programa?

El código "Ethicapp." toma como base los archivos del tipo .csv obtenidos en el sitio web  https://ethicapp.fen.uchile.cl/login y los transforma a data frame (una data para poder trabajarlos en R) y de esta manera poder filtrarlos bajos los requerimientos del usuario, los cuales serían:
Eliminar del estudio las personas que realizan una "etapa 7"
Filtrar a todas las personas que no llevan a cabo una de las 3 etapas (basta con que falte a 1 para sacarlo/a del estudio)
homogenizar la muestra para hacerla comparable (que solo se visualicen 3 etapas: individual 1, grupal e individual 2)
buscar formas para mostrar la información de manera accesible a las personas sin tener el conocimiento a profundidad

Archivos necesarios

ETHICAPP.r: Este es el código en R, no hay más explicación.
Casos.xlsx: Este archivo tiene todos los elementos "básicos" del estudio, cuenta con dos hojas las cuales son "Diferenciales", en donde se guarda el número del diferencial y su descripción con palabras junto con el año en que se realizó, el curso y el nombre del caso. En la segunda hoja, llamada "Cant_est" tenemos los datos de los cursos que fueron impartidos con el semestre, el curso, cada sección y el total de estudiantes. Este archivo es vital si es que se quieren hacer otros casos en el futuro.
Archivos del tipo "año.curso.sec.xlsx" (ejemplo "2022.CD1100.10.xlsx"): estos son los archivos con la información en bruto extraídos de ETHICAPP, no es necesario saber que contienen pero, si es que se requiere hacer estadística inferencial, se puede revisar para poder ver parámetros a utilizar.

