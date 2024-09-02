rm(list=ls())
library(readxl)
library(xlsx)
library(tidyr)
library(dplyr)
library(ggplot2)
library(RColorBrewer)

Cambios<-function(trayectoria){
  a<-switch (trayectoria,
             '1-1' = 'Mantiene',
             '2-1' = 'Radicaliza propia postura 1 i',
             '3-1' = 'Radicaliza propia postura 2 i',
             '4-1' = 'Cambio radical 3 i',
             '5-1' = 'Cambio radical 2 i',
             '6-1' = 'Cambia radical 1 i ',
             '1-2' = 'Modera 1 d',
             '2-2' = 'Mantiene',
             '3-2' = 'Radicaliza moderadamente i',
             '4-2' = 'Moderado a radical i',
             '5-2' = 'Cambio a radical i',
             '6-2' = 'Cambio fuerte a radical i',
             '1-3' = 'modera 2 d',
             '2-3' = 'Modera propia postura i',
             '3-3' = 'Mantiene',
             '4-3' = 'Cambia moderado a moderado i',
             '5-3' = 'Cambia a moderado i',
             '6-3' = 'Cambia fuerte a moderado i',
             '1-4' = 'Cambia fuerte a moderado d',
             '2-4' = 'Cambia a moderado d',
             '3-4' = 'Cambia moderado a moderado d',
             '4-4' = 'Mantiene',
             '5-4' = 'Modera propia postura i',
             '6-4' = 'modera 2 i',
             '1-5' = 'Cambio fuerte a radical d',
             '2-5' = 'Cambio a radical d',
             '3-5' = 'Moderado a radical d',
             '4-5' = 'Radicaliza moderadamente d',
             '5-5' = 'Mantiene',
             '6-5' = 'Modera 1 i',
             '1-6' = 'Cambia radical 1 d',
             '2-6' = 'Cambio radical 2 d',
             '3-6' = 'Cambio radical 3 d',
             '4-6' = 'Radicaliza propia postura 2 d',
             '5-6' = 'Radicaliza propia postura 1 d',
             '6-6' = 'Mantiene'
  )
  return(a)
}

linea<-function(sheet, rowIndex, title, titleStyle,colIndex=1){
  rows <- createRow(sheet, rowIndex=rowIndex)
  if(length(title)==1){
    sheetTitle <- createCell(rows, colIndex=colIndex)
    setCellValue(sheetTitle[[1,1]], title)
    setCellStyle(sheetTitle[[1,1]], titleStyle)
  }else{
    sheetTitle <- createCell(rows, colIndex=c(1:length(title)))
    for(i in c(1:length(title))){
      setCellValue(sheetTitle[[1,i]], title[[i]])
      setCellStyle(sheetTitle[[1,i]], titleStyle[[i]])
    }
  }
}

exportar_excel <- function(parametros,archivo){
  agno<-parametros[[1]]
  cur<-parametros[[2]]
  sec<-parametros[[3]]
  ndiferencial<-parametros[[4]]
  Participacion<-parametros[[5]]
  Cantidad_tabla_etapa<-parametros[[6]]
  Porc_tabla_etapa<-parametros[[7]]
  Tabla_cambio_postura<-parametros[[8]]
  Cant_tablas_magnitud<-parametros[[9]]
  Porc_tablas_magnitud<-parametros[[10]]
  Tabla_direcciones<-parametros[[11]]
  Tabla_mapa_calor<-parametros[[12]]
  data<-parametros[[13]]
  Caso<-parametros[[14]]
  sin_siete<-parametros[[15]]
  
  wb <- createWorkbook(type="xlsx")
  
  
  # Estilos de celdas
  # Estilos de titulos y subtitulos
  titulo <- CellStyle(wb)+ Font(wb,  heightInPoints=16, isBold=TRUE)
  subtitulo <- CellStyle(wb) + Font(wb,  heightInPoints=12,
                                    isItalic=TRUE, isBold=FALSE)
  subtitulo_wraped<- subtitulo+ Alignment(wrapText=TRUE) 
  
  negrita <- CellStyle(wb) + Font(wb, isBold=TRUE)+
    Border(color="black", position=c("TOP"),
           pen=c("BORDER_THICK"))
  
  # Estilo de tablas
  filas_centradas <- CellStyle(wb) + Alignment(wrapText=TRUE) 
  
  columnas <- CellStyle(wb) + Font(wb, isBold=TRUE) +
    Alignment(vertical="VERTICAL_CENTER",wrapText=TRUE, horizontal="ALIGN_CENTER") +
    Border(color="black", position=c("TOP", "BOTTOM"),
           pen=c("BORDER_THICK", "BORDER_THICK"))+Fill(foregroundColor = "lightblue", pattern = "SOLID_FOREGROUND")
  
  sheet <- createSheet(wb, sheetName = "Contexto")
  linea(sheet,1,list("Curso:", paste0(cur,"-",sec)), list(titulo,subtitulo))
  linea(sheet,2,list("Año:", agno), list(titulo,subtitulo))
  linea(sheet,3,list("Cantidad de estudiantes:", unique(Participacion$Total)), list(titulo,subtitulo))
  
  linea(sheet,4,list("Caso:", unique(Caso$Caso)), list(titulo,subtitulo))
  linea(sheet,5,list("Contexto: ", unique(Caso$Contexto)), list(titulo, subtitulo_wraped))
  addDataFrame(Caso%>%select(Diferencial,Pregunta,Izquierda,Derecha)%>%as.data.frame(),
               sheet, startRow=7, startColumn=1,
               colnamesStyle = columnas,
               colStyle = list("1"=filas_centradas,"2"=filas_centradas,"3"=filas_centradas,"4"=filas_centradas),
               row.names = F, characterNA="0")  
  
  setColumnWidth(sheet, colIndex=c(2), colWidth=60)
  setColumnWidth(sheet, colIndex=c(1,3,4), colWidth=32)
  
  
  
  sheet <- createSheet(wb, sheetName = "Recuento")
  if(sin_siete){
    i<-1
  }else{
    linea(sheet,1,"En esta sección no se analizó el cambio de postura debido a que los diferenciales se midieron de 1 a 7",subtitulo)
    i<-3
  }
  
  addDataFrame(Participacion,
               sheet, startRow=i, startColumn=1,
               colnamesStyle = columnas,
               row.names = F, characterNA="0")
  
  fila<-nrow(Participacion)+3+i
  
  addPicture(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Estudiantes_etapa_",agno,cur,sec,".png"), sheet, scale=0.5, startRow=fila , startColumn = 1)
  addPicture(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Box_plot_estudiantes_etapa_",agno,cur,sec,".png"), sheet, scale=0.5, startRow=fila, startColumn = 8*ndiferencial)
  fila<-fila+21
  addPicture(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Magnitud_etapa_",agno,cur,sec,".png"), sheet, scale=0.5, startRow=fila, startColumn = 1)
  if(sin_siete){
    addPicture(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Total_cambio_postura_",agno,cur,sec,".png"), sheet, scale=0.5, startRow=fila, startColumn = 8*ndiferencial)
  }
  fila<-fila+21
  addPicture(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Mapa_calor_transiciones__",agno,cur,sec,".png"), sheet, scale=0.5, startRow=fila, startColumn = 1)
  if(sin_siete){
    addPicture(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Nivel_etapa_",agno,cur,sec,".png"), sheet, scale=0.5, startRow=fila, startColumn = 8*ndiferencial)
  }
  
  setColumnWidth(sheet, colIndex=c(1), colWidth=29)
  setColumnWidth(sheet, colIndex=c(2), colWidth=13)
  
  sheet <- createSheet(wb, sheetName = "Tablas_procesadas")
  
  if(sin_siete){
    i<-1
  }else{
    linea(sheet,1,"En esta sección no se analizó el cambio de postura debido a que los diferenciales se midieron de 1 a 7",subtitulo)
    i<-3
  }
  linea(sheet,rowIndex = i,"Cantidad de estudiantes por opción escogida en cada etapa",titulo)

  i<-2+i
  addDataFrame(subset(Cantidad_tabla_etapa,Diferencial=="Diferencial 1"),
               sheet, startRow=i, startColumn=1,
               colnamesStyle = columnas,
               row.names = F, characterNA="0")
  if(ndiferencial==2){
    addDataFrame(subset(Cantidad_tabla_etapa,Diferencial=="Diferencial 2"),
                 sheet, startRow=i, startColumn=8,
                 colnamesStyle = columnas,
                 row.names = F, characterNA="0")
    i<-i+2+nrow(Cantidad_tabla_etapa)/2
  }else{
    i<-i+2+nrow(Cantidad_tabla_etapa)
  }
  
  linea(sheet,rowIndex = i,"Porcentaje de estudiantes por opción escogida en cada etapa",titulo)
  i<-i+1
  addDataFrame(subset(Porc_tabla_etapa,Diferencial=="Diferencial 1"),
               sheet, startRow=i, startColumn=1,
               colnamesStyle = columnas,
               row.names = F, characterNA="0")
  
  if(ndiferencial==2){
    addDataFrame(subset(Porc_tabla_etapa,Diferencial=="Diferencial 2"),
                 sheet, startRow=i, startColumn=8,
                 colnamesStyle = columnas,
                 row.names = F, characterNA="0")
    i<-i+2+nrow(Porc_tabla_etapa)/2
  }else{
    i<-i+2+nrow(Porc_tabla_etapa)
    
  }
  
    
  if(sin_siete){
    linea(sheet,rowIndex = i,"Cantidad y porcentaje de estudiantes que manifestaron un cambio de postura entre dos etapas",titulo)
    i<-i+1
    addDataFrame(subset(Tabla_cambio_postura,Diferencial=="Diferencial 1"),
                 sheet, startRow=i, startColumn=1,
                 colnamesStyle = columnas,
                 row.names = F, characterNA="0")
    if(ndiferencial==2){
      addDataFrame(subset(Tabla_cambio_postura,Diferencial=="Diferencial 2"),
                   sheet, startRow=i, startColumn=8,
                   colnamesStyle = columnas,
                   row.names = F, characterNA="0")
      i<-i+2+nrow(Tabla_cambio_postura)/2
    }else{
      i<-i+2+nrow(Tabla_cambio_postura)
    }
  }
  
  
  linea(sheet,rowIndex = i,"Cantidad de estudiantes según magnitud de la variación entre dos etapas",titulo)
  i<-i+1
  addDataFrame(subset(Cant_tablas_magnitud,Diferencial=="Diferencial 1"),
               sheet, startRow=i, startColumn=1,
               colnamesStyle = columnas,
               row.names = F, characterNA="0")
  if(ndiferencial==2){
    addDataFrame(subset(Cant_tablas_magnitud,Diferencial=="Diferencial 2"),
                 sheet, startRow=i, startColumn=8,
                 colnamesStyle = columnas,
                 row.names = F, characterNA="0")
    i<-i+2+nrow(Cant_tablas_magnitud)/2
  }else{
    i<-i+2+nrow(Cant_tablas_magnitud)
    
  }
  
  
  linea(sheet,rowIndex = i,"Porcentaje de estudiantes según magnitud de la variación entre dos etapas",titulo)
  i<-i+1
  addDataFrame(subset(Porc_tablas_magnitud,Diferencial=="Diferencial 1"),
               sheet, startRow=i, startColumn=1,
               colnamesStyle = columnas,
               row.names = F, characterNA="0")
  if(ndiferencial==2){
    addDataFrame(subset(Porc_tablas_magnitud,Diferencial=="Diferencial 1"),
                 sheet, startRow=i, startColumn=8,
                 colnamesStyle = columnas,
                 row.names = F, characterNA="0")
    i<-i+2+nrow(Porc_tablas_magnitud)/2
  }else{
    i<-i+2+nrow(Porc_tablas_magnitud)
    
  }
  
  
  if(sin_siete){
    linea(sheet,rowIndex = i,"Nivel y dirección en la que se produjo la variación entre dos etapas",titulo)
    i<-i+1
    addDataFrame(subset(Tabla_direcciones,Diferencial=="Diferencial 1"),
                 sheet, startRow=i, startColumn=1,
                 colnamesStyle = columnas,
                 row.names = F, characterNA="0")
    if(ndiferencial==2){
      addDataFrame(subset(Tabla_direcciones,Diferencial=="Diferencial 2"),
                   sheet, startRow=i, startColumn=8,
                   colnamesStyle = columnas,
                   row.names = F, characterNA="0")
      i<-i+2+max(nrow(Tabla_direcciones%>%filter(Diferencial=="Diferencial 1")),nrow(Tabla_direcciones%>%filter(Diferencial=="Diferencial 2")))
    }else{
      i<-i+2+nrow(Tabla_direcciones)
    }
  }
  
  
  linea(sheet,rowIndex = i,"Cantidad de estudiantes según opción original y final en cada etapa",titulo)
  i<-i+1
  addDataFrame(subset(Tabla_mapa_calor,Diferencial=="Diferencial 1"),
               sheet, startRow=i, startColumn=1,
               colnamesStyle = columnas,
               row.names = F, characterNA="0")
  if(ndiferencial==2){
    addDataFrame(subset(Tabla_mapa_calor,Diferencial=="Diferencial 2"),
                 sheet, startRow=i, startColumn=8,
                 colnamesStyle = columnas,
                 row.names = F, characterNA="0")
    i<-i+2+nrow(Tabla_mapa_calor)/2
  }else{
    i<-i+2+nrow(Tabla_mapa_calor)
  }
  
  setColumnWidth(sheet, colIndex=c(1,3,8,10), colWidth=17)
  setColumnWidth(sheet, colIndex=c(2,9), colWidth=15)
  setColumnWidth(sheet, colIndex=c(4,11), colWidth=25)
  
  
  sheet <- createSheet(wb, sheetName = "Datos")
  addDataFrame(data,
               sheet, startRow=1, startColumn=1,
               colnamesStyle = columnas,
               row.names = F, characterNA="0")
  
  # Guardar
  saveWorkbook(wb, archivo)
}

analizar<-function(agno,cur,secciones= c(1:9,"RZ","RZ2","RZ3","Todas_las_sec")){
  Casos<-read_excel("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\Casos.xlsx",sheet="Diferenciales")
  Cant_est<-read_excel("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\Casos.xlsx","Cant_est")
  rm(Todos_los_cursos)
  for (sec in secciones){
    print(sec)
    if(!sec=="Todas_las_sec"){
      sem<-ifelse(cur=="CD1100",1,2)
      file<-paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\",agno,".",cur,".",sec,".csv")
      continuar<-file.exists(file)
    }else{
      continuar<-T
    }
    if(continuar){
      Caso<-Casos%>%filter(Agno==agno,Curso==cur)
      ndiferencial<-length(Caso$Diferencial)
      length(Caso$Diferencial)
      if(!sec=="Todas_las_sec"){
        data<-read.csv(file,encoding = "UTF-8",sep=";")
        data<-data%>%select(rut,
                            team_id,
                            df,
                            opt_left,
                            sel,
                            phase,
                            user_id,
                            comment)
        data$Diferencial<-data$df
        #data$Diferencial <- ifelse(data$opt_left == "Preservar el recurso natural escaso", 1, 2)
        #cambio para Adela 2021, solo tiene df=1
        data$Diferencial[data$Diferencial==1]<-"Diferencial 1"
        if(ndiferencial>=2){
          data$Diferencial[data$Diferencial==2]<-"Diferencial 2"
          if(ndiferencial==3){
            data$Diferencial[data$Diferencial==3]<-"Diferencial 3"
          }
        }
        if(agno==2021 & cur=="CD1100"){
          data<-
            data%>%
            filter(phase %in% c(3,4,5))
          data$fase<-NA
          data$fase[data$phase==3]<-"Ind1"
          data$fase[data$phase==4]<-"Grup"
          data$fase[data$phase==5]<-"Ind2"
        }else if(agno==2021 & cur=="CD1201"){
          data$fase[data$phase %in% c(1,2)]<-"Ind1"
          data$fase[data$phase %in% c(3,4)]<-"Grup"
          data$fase[data$phase %in% c(5,6)]<-"Ind2"
        }else if(agno==2021 & cur=="CD2201"){
          data<-data%>%filter(phase>1)
          data$fase[data$phase %in% c(2)]<-"Ind1" #data$fase[data$phase %in% c(1,2)]<-"Ind1"
          data$fase[data$phase %in% c(3)]<-"Grup"  #data$fase[data$phase %in% c(3,4)]<-"Grup"
          data$fase[data$phase %in% c(4)]<-"Ind2"  #data$fase[data$phase %in% c(5,6)]<-"Ind2"
        }else if(agno==2022 & cur=="CD1100"){
          data<-data%>%filter(phase>1)
          data$fase[data$phase %in% c(2)]<-"Ind1"
          data$fase[data$phase %in% c(3)]<-"Grup"
          data$fase[data$phase %in% c(4)]<-"Ind2"
        }else if(agno==2022 & cur=="CD1201"){
          data<-data%>%filter(phase>1)
          data$fase[data$phase %in% c(2)]<-"Ind1"
          data$fase[data$phase %in% c(3)]<-"Grup"
          data$fase[data$phase %in% c(4)]<-"Ind2"
        }else if(agno==2023 & cur=="CD1100"){
          data<-data%>%filter(phase>1)
          data$fase[data$phase %in% c(2)]<-"Ind1"
          data$fase[data$phase %in% c(3)]<-"Grup"
          data$fase[data$phase %in% c(4)]<-"Ind2"
        }else if(agno==2023 & cur=="CD1201"){
          data$fase[data$phase %in% c(1)]<-"Ind1"
          data$fase[data$phase %in% c(2)]<-"Grup"
          data$fase[data$phase %in% c(3)]<-"Ind2"
        }else if(agno==2024 & cur=="CD1100"){
          data$fase[data$phase %in% c(1)]<-"Ind1"
          data$fase[data$phase %in% c(2)]<-"Grup"
          data$fase[data$phase %in% c(3)]<-"Ind2"
        }else if(agno==2024 & cur=="EH2202"){
          data$fase[data$phase %in% c(1)]<-"Ind1"
          data$fase[data$phase %in% c(2)]<-"Grup"
          data$fase[data$phase %in% c(3)]<-"Ind2"
        } 
        sin_siete<-(!7 %in% data$sel)
        N_estudiantes<-ifelse(!sec=="RZ",Cant_est%>%filter(Curso==cur,Semestre==agno*10+sem,Seccion==sec)%>%pull(Cantidad),0)
        if(sin_siete){
          if(exists("Todos_los_cursos")){
            Todos_los_cursos<-rbind(Todos_los_cursos,data%>%filter(!rut %in% Todos_los_cursos$rut))
            Total_estudiantes<-Total_estudiantes+N_estudiantes
          }else{
            Todos_los_cursos<-data
            Total_estudiantes<-N_estudiantes
          }
        }
        
      }else{
        data<-Todos_los_cursos
        sin_siete<-T
        N_estudiantes<-Total_estudiantes
      }
      
      otros_datos<-
        data%>%
        select(rut,
               user_id,
               Diferencial,
               fase,
               comment)%>%
        mutate(Fase=paste0("Comentario - ",fase," - ",Diferencial),
               fase=factor(fase,levels=c("Ind1","Grup","Ind2")))%>%
        arrange(fase,Diferencial)%>%
        mutate(Fase=factor(Fase,levels=unique(Fase)))%>%
        select(-Diferencial,-fase)%>%
        spread(Fase,comment)
      data<-
        data%>%
        select(-user_id,
               -comment)
      Cant_resp<-
        data%>%
        group_by(fase,Diferencial)%>%
        count()%>%
        mutate(Total=ifelse(!sec=="RZ",N_estudiantes,NA),
               Porc=ifelse(!sec=="RZ",round(n/Total*100),NA))

      #Homogeniza los textos de los extremos
      data$df <- as.numeric(data$df)
      unique_values_2 <- unique(subset(data$opt_left, data$df == "2"))
      
      data <- data %>%
        mutate(opt_left = ifelse(opt_left!=unique_values_2[1],unique_values_2[1], opt_left))
      
      unique_values_1 <- unique(subset(data$opt_left, data$df == "1"))
      
      data <- data %>%
        mutate(opt_left = ifelse(opt_left!=unique_values_1[1],unique_values_1[1], opt_left))
      
      
      data<-data%>%
        select(-team_id,-phase)%>%
        spread(fase,sel)

      data$Magnitud_Ind1_Grup<-NA
      data$Magnitud_Grup_Ind2<-NA
      data$Magnitud_Ind1_Ind2<-NA
      data<-data[!(is.na(data$Ind1)), ]
      data<-data[!(is.na(data$Ind2)), ]
      data<-data[!(is.na(data$Grup)), ]
      data$Magnitud_Ind1_Grup<-abs(data$Ind1-data$Grup)
      data$Magnitud_Grup_Ind2<-abs(data$Grup-data$Ind2)
      data$Magnitud_Ind1_Ind2<-abs(data$Ind1-data$Ind2)
      if(sin_siete){
        data$Cambio_postura_Ind1_Grup<-((data$Ind1<=3)*(data$Grup>=4))+((data$Ind1>=4)*(data$Grup<=3))
        data$Cambio_postura_Grup_Ind2<-((data$Grup<=3)*(data$Ind2>=4))+((data$Grup>=4)*(data$Ind2<=3))
        data$Cambio_postura_Ind1_Ind2<-((data$Ind1<=3)*(data$Ind2>=4))+((data$Ind1>=4)*(data$Ind2<=3))
        data$Nivel_Ind1_Grup<-vapply(paste0(data$Ind1,"-",data$Grup),Cambios,c("a"))
        data$Nivel_Grup_Ind2<-vapply(paste0(data$Grup,"-",data$Ind2),Cambios,c("a"))
        data$Nivel_Ind1_Ind2<-vapply(paste0(data$Ind1,"-",data$Ind2),Cambios,c("a"))
        data<-data%>%
          mutate(Nivel_Ind1_Grup=ifelse(Magnitud_Ind1_Grup %in% c(4,5),"Cambio extremo",
                                        ifelse(Magnitud_Ind1_Grup %in% c(2,3), "Cambio moderado",
                                               ifelse(Magnitud_Ind1_Grup %in% c(1), "Cambio pequeño",
                                                      "Se mantiene"))),
                 Nivel_Ind1_Ind2=ifelse(Magnitud_Ind1_Ind2 %in% c(4,5),"Cambio extremo",
                                        ifelse(Magnitud_Ind1_Ind2 %in% c(2,3), "Cambio moderado",
                                               ifelse(Magnitud_Ind1_Ind2 %in% c(1), "Cambio pequeño",
                                                      "Se mantiene"))),
                 Nivel_Grup_Ind2=ifelse(Magnitud_Grup_Ind2 %in% c(4,5),"Cambio extremo",
                                        ifelse(Magnitud_Grup_Ind2 %in% c(2,3), "Cambio moderado",
                                               ifelse(Magnitud_Grup_Ind2 %in% c(1), "Cambio pequeño",
                                                      "Se mantiene"))),
                 Direccion_Ind1_Grup=ifelse(Cambio_postura_Ind1_Grup==1 & Grup>=4, "Cambia hacia la derecha",
                                            ifelse(Cambio_postura_Ind1_Grup==0 & Grup>=4, "Se mantiene en la derecha",
                                                   ifelse(Cambio_postura_Ind1_Grup==1 & Grup<=3, "Cambia hacia la izquierda","Se mantiene en la izquierda"))),
                 Direccion_Grup_Ind2=ifelse(Cambio_postura_Grup_Ind2==1 & Ind2>=4, "Cambia hacia la derecha",
                                            ifelse(Cambio_postura_Grup_Ind2==0 & Ind2>=4, "Se mantiene en la derecha",
                                                   ifelse(Cambio_postura_Grup_Ind2==1 & Ind2<=3, "Cambia hacia la izquierda","Se mantiene en la izquierda"))),
                 Direccion_Ind1_Ind2=ifelse(Cambio_postura_Ind1_Ind2==1 & Ind2>=4, "Cambia hacia la derecha",
                                            ifelse(Cambio_postura_Ind1_Ind2==0 & Ind2>=4, "Se mantiene en la derecha",
                                                   ifelse(Cambio_postura_Ind1_Ind2==1 & Ind2<=3, "Cambia hacia la izquierda","Se mantiene en la izquierda")))
          )
      }
      
      n_cant_resp<-nrow(Cant_resp)
      i<-1
      for(df in Caso$Diferencial){
        Cant_resp[n_cant_resp+i,1]<-"Respondieron todas las etapas"
        Cant_resp[n_cant_resp+i,2]<-paste0("Diferencial ",df)
        Cant_resp[n_cant_resp+i,3]<-length(data$rut[data$Diferencial==paste0("Diferencial ",df)])
        print(length(data$rut[data$Diferencial==paste0("Diferencial ",df)]))
        Cant_resp[n_cant_resp+i,4]<-ifelse(!sec=="RZ",N_estudiantes,NA)
        Cant_resp[n_cant_resp+i,5]<-ifelse(!sec=="RZ",round(Cant_resp[n_cant_resp+1,3]/N_estudiantes*100),NA)
        i<-i+1
      }
      
      Cant_resp$fase<-factor(Cant_resp$fase,levels=c("Ind1","Grup","Ind2","Respondieron todas las etapas"))
      Cant_resp$Diferencial<-factor(Cant_resp$Diferencial,levels=c("Diferencial 1","Diferencial 2"))
      Cant_resp<-
        Cant_resp%>%
        arrange(Diferencial)
      data1<- filter(data,Diferencial=="Diferencial 1")
      data2<- data
      
      Tabla_etapa<-
        data2%>%
        select(Diferencial,Ind1,Grup,Ind2)%>%
        gather(Etapa,Valor,Ind1:Ind2)%>%
        mutate(Etapa=factor(Etapa,levels=c("Ind1","Grup","Ind2")))%>%
        group_by(Diferencial,Etapa,Valor)%>%
        count()%>%
        ungroup()%>%
        group_by(Etapa,Diferencial)%>%
        mutate(Porc=n/sum(n)*100)
      
      p<-
        Tabla_etapa%>%
        ggplot(aes(x=Etapa,y=Porc,fill=forcats::fct_rev(factor(Valor))))+
        geom_bar(stat="identity",position = "stack")+
        geom_text(aes(label=round(Porc)),colour="black", position = position_stack(vjust=0.5),show.legend = F,size = 5)+
        ggtitle("Porcentaje de estudiantes que seleccionaron opción según etapa")+
        theme(plot.title = element_text(hjust = 0.5,size = 20),
              axis.title = element_text(size = 22),
              axis.text=element_text(size = 22),
              strip.text = element_text(size = 22),
              legend.text = element_text(size = 22),
              legend.title = element_text(size = 22))+
        facet_wrap(.~Diferencial)
      paleta<-c()
      if(sin_siete){
        if(6 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,1)
        }
        if(5 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,2)
        }
        if(4 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,3)
        }
        if(3 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,9)
        }
        if(2 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,10)
        }
        if(1 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,11)
        }
      }else{
        if(7 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,1)
        }
        if(6 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,2)
        }
        if(5 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,3)
        }
        if(4 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,6)
        }
        if(3 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,9)
        }
        if(2 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,10)
        }
        if(1 %in% Tabla_etapa$Valor){
          paleta<-c(paleta,11)
        }
      }
      p<-p+scale_fill_manual(name="Valor",values = brewer.pal(11,"RdYlGn")[paleta]) 
      ggsave(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Estudiantes_etapa_",agno,cur,sec,".png"),plot=p,height=20,width=20*ndiferencial,dpi=300,units="cm",bg="white")  
      
      Cantidad_tabla_etapa<-
        Tabla_etapa%>%
        select(-Porc)%>%
        spread(Etapa,n)%>%
        rename(Opción = Valor)
      
      Porc_tabla_etapa<-
        Tabla_etapa%>%
        select(-n)%>%
        spread(Etapa,Porc)%>%
        rename(Opción = Valor)
      
      
      p<-
        data2%>%
        select(Diferencial,Ind1,Grup,Ind2)%>%
        gather(Etapa,Valor,Ind1:Ind2)%>%
        mutate(Etapa=factor(Etapa,levels=c("Ind1","Grup","Ind2")))%>%
        ggplot(aes(x=Etapa,y=Valor,fill=Etapa))+
        geom_boxplot()+
        ggtitle("Boxplots de la opción seleccionada según etapa")+
        facet_wrap(.~Diferencial)+
        theme(plot.title = element_text(hjust = 0.5,size = 20),
              axis.title = element_text(size = 22),
              axis.text=element_text(size = 22),
              strip.text = element_text(size = 22),
              legend.position = "none")+
        scale_fill_manual(values=c("#1F4364","#3A78AA","#56B1F7"))
      if(sin_siete){
        p<-p+
          scale_y_continuous(breaks=c(1:6))
      }else{
        p<-p+
          scale_y_continuous(breaks=c(1:7))
      }
      ggsave(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Box_plot_estudiantes_etapa_",agno,cur,sec,".png"),plot=p,height=20,width=20*ndiferencial,dpi=300,units="cm",bg="white")  
      if(sin_siete){
        Tabla_cambio_postura<-
          data2%>%
          select(Diferencial,Cambio_postura_Ind1_Grup,Cambio_postura_Grup_Ind2,Cambio_postura_Ind1_Ind2)%>%
          gather(Etapas,Valor,Cambio_postura_Ind1_Grup:Cambio_postura_Ind1_Ind2)%>%
          mutate(Etapas=gsub("Cambio_postura_","",Etapas),
                 Etapas=factor(Etapas,levels=c("Ind1_Grup","Grup_Ind2","Ind1_Ind2")))%>%
          group_by(Diferencial,Etapas)%>%
          summarise(n=sum(Valor))%>%
          ungroup()%>%
          mutate(Porc=n/Cant_resp$n[Cant_resp$fase=="Respondieron todas las etapas"]*100)  
        
        p<-
          Tabla_cambio_postura%>%
          ggplot(aes(x=Etapas,y=Porc,fill=Etapas))+
          geom_bar(stat="identity")+
          geom_text(aes(label=round(Porc)),colour="black", vjust=3,show.legend = F,size = 5)+
          scale_fill_manual(values=c("#1F4364","#3A78AA","#56B1F7"))+
          ggtitle("Porcentaje de estudiantes que cambiaron de postura según par de etapas")+
          theme(plot.title = element_text(hjust = 0.5,size = 20),
                axis.title = element_text(size = 22),
                axis.text=element_text(size = 22),
                strip.text = element_text(size = 22),
                legend.position = "none")+
          facet_wrap(.~Diferencial)
        ggsave(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Total_cambio_postura_",agno,cur,sec,".png"),plot=p,height=20,width=20*ndiferencial,dpi=300,units="cm",bg="white")
      }else{
        Tabla_cambio_postura<-NA
      }

      Tablas_magnitud<-
        data2%>%
        select(Diferencial,Magnitud_Ind1_Grup,Magnitud_Grup_Ind2,Magnitud_Ind1_Ind2)%>%
        gather(Etapas,Valor,Magnitud_Ind1_Grup:Magnitud_Ind1_Ind2)%>%
        mutate(Etapas=gsub("Magnitud_","",Etapas),
               Etapas=factor(Etapas,levels=c("Ind1_Grup","Grup_Ind2","Ind1_Ind2")))%>%
        group_by(Diferencial,Etapas,Valor)%>%
        count()%>%
        ungroup()%>%
        group_by(Diferencial,Etapas)%>%
        mutate(Porc=n/sum(n)*100)
      
      p<-
        Tablas_magnitud%>%
        ggplot(aes(x=Etapas,y=Porc,fill=forcats::fct_rev(factor(Valor))))+
        geom_bar(stat="identity",position = "stack")+
        geom_text(aes(label=round(Porc)),colour="black", position = position_stack(vjust=0.5),show.legend = F,size = 5)+
        ggtitle("Porcentaje de estudiantes según magnitud de la variación entre cada par de etapas")+
        theme(plot.title = element_text(hjust = 0.5,size = 20),
              axis.title = element_text(size = 22),
              axis.text=element_text(size = 22),
              strip.text = element_text(size = 22),
              legend.text = element_text(size = 22),
              legend.title = element_text(size = 22))+
        facet_wrap(.~Diferencial)
      paleta<-c()
      if(sin_siete){
        if(5 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,1)
        }
        if(4 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,2)
        }
        if(3 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,3)
        }
        if(2 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,9)
        }
        if(1 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,10)
        }
        if(0 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,11)
        }
      }else{
        if(6 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,1)
        }
        if(5 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,2)
        }
        if(4 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,3)
        }
        if(3 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,6)
        }
        if(2 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,9)
        }
        if(1 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,10)
        }
        if(0 %in% Tablas_magnitud$Valor){
          paleta<-c(paleta,11)
        }
      }
      p<-p+scale_fill_manual(name="Valor",values = brewer.pal(11,"RdYlGn")[paleta]) 
      ggsave(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Magnitud_etapa_",agno,cur,sec,".png"),plot=p,height=20,width=20*ndiferencial,dpi=300,units="cm",bg="white")
      
      Tablas_magnitud<-
        Tablas_magnitud%>%
        rename(Magnitud=Valor)
      
      Cant_tablas_magnitud<-
        Tablas_magnitud%>%
        select(-Porc)%>%
        spread(Etapas,n)
      Cant_tablas_magnitud[is.na(Cant_tablas_magnitud)]<-0
      
      Porc_tablas_magnitud<-
        Tablas_magnitud%>%
        select(-n)%>%
        spread(Etapas,Porc)
      Cant_tablas_magnitud[is.na(Cant_tablas_magnitud)]<-0
      
      if(sin_siete){
        Tabla_direcciones<-
          data%>%
          select(Diferencial,Nivel_Ind1_Grup,Nivel_Grup_Ind2,Nivel_Ind1_Ind2,Direccion_Ind1_Grup,Direccion_Grup_Ind2,Direccion_Ind1_Ind2)%>%
          gather(Direccion_etapas,Direccion,Direccion_Ind1_Grup:Direccion_Ind1_Ind2)%>%
          gather(Nivel_etapas,Nivel,Nivel_Ind1_Grup:Nivel_Ind1_Ind2)%>%
          filter(gsub("Nivel_","",Nivel_etapas)==gsub("Direccion_","",Direccion_etapas))%>%
          mutate(Etapas=gsub("Direccion_","",Direccion_etapas),
                 Nivel_etapas=factor(Nivel_etapas,levels=c("Nivel_Ind1_Grup","Nivel_Grup_Ind2","Nivel_Ind1_Ind2")),
                 Direccion_etapas=factor(Direccion_etapas,levels=c("Direccion_Ind1_Grup","Direccion_Grup_Ind2","Direccion_Ind1_Ind2")),
          )%>%
          group_by(Diferencial,Etapas,Nivel,Direccion)%>%
          count()%>%
          ungroup()%>%
          group_by(Diferencial,Etapas)%>%
          mutate(Porc=n/sum(n)*100)
        
        if(ndiferencial==1){
          p<-
            Tabla_direcciones%>%
            ggplot(aes(x=Etapas,y=n,fill=Nivel))+
            geom_bar(stat="identity",position = "stack")+
            geom_text(aes(label=round(n)),colour="black", position = position_stack(vjust=0.5),show.legend = F,size = 5)+
            ggtitle("Cantidad de estudiantes según nivel y dirección de la variación entre cada par de etapas")+
            theme(plot.title = element_text(hjust = 0.5,size = 10),
                  axis.title = element_text(size = 10),
                  axis.text=element_text(size = 10),
                  strip.text = element_text(size = 10),
                  legend.text = element_text(size = 10),
                  legend.title = element_blank())+
            facet_wrap(.~Direccion)+
            scale_fill_manual(name="Tipo cambio",values = brewer.pal(11,"RdYlGn")[c(1,2,10,11)]) 
        }else{
          p1<-
            Tabla_direcciones%>%
            filter(Diferencial=="Diferencial 1")%>%
            ggplot(aes(x=Etapas,y=n,fill=Nivel))+
            geom_bar(stat="identity",position = "stack")+
            geom_text(aes(label=round(n)),colour="black", position = position_stack(vjust=0.5),show.legend = F,size = 5)+
            ggtitle("Diferencial 1")+
            theme(plot.title = element_text(hjust = 0.5,size = 10),
                  axis.title = element_text(size = 10),
                  axis.text=element_text(size = 10),
                  strip.text = element_text(size = 10),
                  legend.position = "none")+
            facet_wrap(.~Direccion)+
            scale_fill_manual(name="Tipo cambio",values = brewer.pal(11,"RdYlGn")[c(1,2,10,11)]) 
          
          p2<-
            Tabla_direcciones%>%
            filter(Diferencial=="Diferencial 2")%>%
            ggplot(aes(x=Etapas,y=n,fill=Nivel))+
            geom_bar(stat="identity",position = "stack")+
            geom_text(aes(label=round(n)),colour="white", position = position_stack(vjust=0.5),show.legend = F,size = 5)+
            ggtitle("Diferencial 2")+
            theme(plot.title = element_text(hjust = 0.5,size = 10),
                  axis.title = element_text(size = 10),
                  axis.text=element_text(size = 10),
                  strip.text = element_text(size = 10),
                  legend.text = element_text(size = 10),
                  legend.title = element_text(size = 10))+
            facet_wrap(.~Direccion)+
            theme(legend.title = element_blank())+
            scale_fill_manual(name="Tipo cambio",values = brewer.pal(11,"RdYlGn")[c(1,2,10,11)]) 
          p<-egg::ggarrange(p1,p2,ncol = 2)
          p<-ggpubr::annotate_figure(p, top = ggpubr::text_grob("Cantidad de estudiantes según nivel y dirección de la variación entre cada par de etapas",size = 14))
        }
        ggsave(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Nivel_etapa_",agno,cur,sec,".png"),plot=p,height=20,width=20*ndiferencial,dpi=300,units="cm",bg="white")
      }else{
        Tabla_direcciones<-NA
      }
      
      if(sin_siete){
        N_max<-6
      }else{
        N_max<-7
      }
      tabla_Ind1_Grup<-
        data%>%
        group_by(Diferencial,Ind1,Grup)%>%
        count()%>%
        ungroup()%>%
        mutate(Ind1=factor(Ind1,levels=c(1:N_max)),
               Grup=factor(Grup,levels=c(1:N_max)))%>%
        complete(Diferencial,Ind1,Grup,fill=list(n=0))%>%
        mutate(Ind1=as.numeric(Ind1),
               Grup=as.numeric(Grup))
      
      tabla_Grup_Ind2<-
        data%>%
        group_by(Diferencial,Grup,Ind2)%>%
        count()%>%
        ungroup()%>%
        mutate(Ind2=factor(Ind2,levels=c(1:N_max)),
               Grup=factor(Grup,levels=c(1:N_max)))%>%
        complete(Diferencial,Grup,Ind2,fill=list(n=0))%>%
        mutate(Grup=as.numeric(Grup),
               Ind2=as.numeric(Ind2))
      
      tabla_Ind1_Ind2<-
        data%>%
        group_by(Diferencial,Ind1,Ind2)%>%
        count()%>%
        ungroup()%>%
        mutate(Ind1=factor(Ind1,levels=c(1:N_max)),
               Ind2=factor(Ind2,levels=c(1:N_max)))%>%
        complete(Diferencial,Ind1,Ind2,fill=list(n=0))%>%
        mutate(Ind1=as.numeric(Ind1),
               Ind2=as.numeric(Ind2))
      
      if(ndiferencial==1){
        p1<-
          tabla_Ind1_Grup%>%
          ggplot(aes(x=Ind1,y=Grup,fill=n))+
          geom_tile(color="black",lwd=1,linetype=1)+
          scale_fill_gradient(low = "white", high = "red")+
          geom_text(aes(label = n), size = 4) +
          theme(legend.position = "none")+
          ggtitle("Ind1 a Grup")+
          theme(plot.title = element_text(hjust = 0.5,size = 13),
                axis.title = element_text(size = 15),
                axis.text=element_text(size = 15),
                strip.text = element_text(size = 15))+
          scale_x_continuous(breaks = c(1:N_max))+
          scale_y_continuous(breaks = c(1:N_max))
        p2<-
          tabla_Grup_Ind2%>%
          ggplot(aes(x=Grup,y=Ind2,fill=n))+
          geom_tile(color="black",lwd=1,linetype=1)+
          scale_fill_gradient(low = "white", high = "red")+
          geom_text(aes(label = n), size = 4) +
          theme(legend.position = "none")+
          ggtitle("Grup a Ind2")+
          theme(plot.title = element_text(hjust = 0.5,size = 13),
                axis.title = element_text(size = 15),
                axis.text=element_text(size = 15),
                strip.text = element_text(size = 15))+
          scale_x_continuous(breaks = c(1:N_max))+
          scale_y_continuous(breaks = c(1:N_max))
        p3<-
          tabla_Ind1_Ind2%>%
          ggplot(aes(x=Ind1,y=Ind2,fill=n))+
          geom_tile(color="black",lwd=1,linetype=1)+
          scale_fill_gradient(low = "white", high = "red")+
          geom_text(aes(label = n), size = 4) +
          ggtitle("Ind1 a Ind2")+
          theme(plot.title = element_text(hjust = 0.5,size = 13),
                axis.title = element_text(size = 15),
                axis.text=element_text(size = 15),
                strip.text = element_text(size = 15))+
          theme(legend.position = "none")+
          scale_x_continuous(breaks = c(1:N_max))+
          scale_y_continuous(breaks = c(1:N_max))
        p<- ggpubr::ggarrange(p1,p2,p3,ncol = 3)
        p<-ggpubr::annotate_figure(p,top = ggpubr::text_grob("Cantidad de estudiantes según transiciones",size = 14),
                                   bottom = ggpubr::text_grob("Etapa inicial",size = 14),
                                   left = ggpubr::text_grob("Etapa final",size = 14, rot = 90))
      }else{
        p<-list()
        for(df in unique(tabla_Ind1_Grup$Diferencial)){
          p1<-
            tabla_Ind1_Grup%>%
            filter(Diferencial==df)%>%
            ggplot(aes(x=Ind1,y=Grup,fill=n))+
            geom_tile(color="black",lwd=1,linetype=1)+
            scale_fill_gradient(low = "white", high = "red")+
            geom_text(aes(label = n), size = 4) +
            theme(legend.position = "none")+
            ggtitle("Ind1 a Grup")+
            theme(plot.title = element_text(hjust = 0.5,size = 13),
                  axis.title = element_text(size = 15),
                  axis.text=element_text(size = 15),
                  strip.text = element_text(size = 15))+
            scale_x_continuous(breaks = c(1:N_max))+
            scale_y_continuous(breaks = c(1:N_max))             
          p2<-
            tabla_Grup_Ind2%>%
            filter(Diferencial==df)%>%
            ggplot(aes(x=Grup,y=Ind2,fill=n))+
            geom_tile(color="black",lwd=1,linetype=1)+
            scale_fill_gradient(low = "white", high = "red")+
            geom_text(aes(label = n), size = 4) +
            theme(legend.position = "none")+
            ggtitle("Grup a Ind2")+
            theme(plot.title = element_text(hjust = 0.5,size = 13),
                  axis.title = element_text(size = 15),
                  axis.text=element_text(size = 15),
                  strip.text = element_text(size = 15))+
            scale_x_continuous(breaks = c(1:N_max))+
            scale_y_continuous(breaks = c(1:N_max))             
          p3<-
            tabla_Ind1_Ind2%>%
            filter(Diferencial==df)%>%
            ggplot(aes(x=Ind1,y=Ind2,fill=n))+
            geom_tile(color="black",lwd=1,linetype=1)+
            scale_fill_gradient(low = "white", high = "red")+
            geom_text(aes(label = n), size = 4) +
            theme(legend.position = "none")+
            ggtitle("Ind1 a Ind2")+
            theme(plot.title = element_text(hjust = 0.5,size = 13),
                  axis.title = element_text(size = 15),
                  axis.text=element_text(size = 15),
                  strip.text = element_text(size = 15))+
            theme(legend.position = "none")+
            scale_x_continuous(breaks = c(1:N_max))+
            scale_y_continuous(breaks = c(1:N_max))
          
          p[[df]]<-egg::ggarrange(p1,p2,p3,ncol = 3)
          p[[df]]<-ggpubr::annotate_figure(p[[df]], top = ggpubr::text_grob(df,size = 10))
        }
        
        if(ndiferencial==2){
          p<-egg::ggarrange(p[["Diferencial 1"]],p[["Diferencial 2"]],ncol = 2)
        }else{
          p<-egg::ggarrange(p[["Diferencial 1"]],p[["Diferencial 2"]],p[["Diferencial 3"]],ncol = 3)
        }
        p<-ggpubr::annotate_figure(p, top = ggpubr::text_grob("Cantidad de estudiantes según transiciones",size = 14),
                                   bottom = ggpubr::text_grob("Etapa inicial",size = 14),
                                   left = ggpubr::text_grob("Etapa final",size = 14,rot=90))
      }
      ggsave(paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\ETHICAPP\\Mapa_calor_transiciones__",agno,cur,sec,".png"),plot=p,height=20,width=20*ndiferencial,dpi=300,units="cm",bg="white")
      tabla_Ind1_Grup<-
        tabla_Ind1_Grup%>%
        mutate(Etapas="Ind1_Grup")%>%
        rename(Opción_original=Ind1,
               Opción_final=Grup)
      
      tabla_Grup_Ind2<-
        tabla_Grup_Ind2%>%
        mutate(Etapas="Grup_Ind2")%>%
        rename(Opción_original=Grup,
               Opción_final=Ind2)
      
      tabla_Ind1_Ind2<-
        tabla_Ind1_Ind2%>%
        mutate(Etapas="Ind1_Ind2")%>%
        rename(Opción_original=Ind1,
               Opción_final=Ind2)
      
      Tabla_mapa_calor<-
        rbind(
          tabla_Ind1_Grup,
          tabla_Grup_Ind2,
          tabla_Ind1_Ind2
        )
      
      data<-
        merge(data,
              otros_datos,
              by="rut")
      data<-data[,c("rut","user_id",colnames(data)[!colnames(data) %in% c("rut","user_id")])]
      
      uniformizar<-function(data_,rellenar=T){
        if(prod(is.na(data_))==0){
          data_<-
            data_%>%
            mutate_if(is.numeric,round)%>%
            as.data.frame()
          if(rellenar){
            data_[is.na(data_)]<-0
          }
          return(data_)
        }else{
          return(NA)
        }
      }
      Cant_resp<-uniformizar(Cant_resp,rellenar = F)
      Cantidad_tabla_etapa<-uniformizar(Cantidad_tabla_etapa)
      Porc_tabla_etapa<-uniformizar(Porc_tabla_etapa)
      Tabla_cambio_postura<-uniformizar(Tabla_cambio_postura)
      Cant_tablas_magnitud<-uniformizar(Cant_tablas_magnitud)
      Porc_tablas_magnitud<-uniformizar(Porc_tablas_magnitud)
      Tabla_direcciones<-uniformizar(Tabla_direcciones)
      Tabla_mapa_calor<-uniformizar(Tabla_mapa_calor)
      
      parametros<-
        list(agno,
             cur,
             sec,
             ndiferencial,
             Cant_resp,
             Cantidad_tabla_etapa,
             Porc_tabla_etapa,
             Tabla_cambio_postura,
             Cant_tablas_magnitud,
             Porc_tablas_magnitud,
             Tabla_direcciones,
             Tabla_mapa_calor,
             data,
             Caso,
             sin_siete)
      exportar_excel(parametros,paste0("C:\\Users\\hurmi\\OneDrive\\Escritorio\\u\\practica 1\\Resultados por sección CD1100 y CD1201 2021\\Procesamient_",agno,"_",cur,"_",sec,".xlsx"))
      
    }
  }
}
analizar(2021,"CD1100",c(1:9,"RZ","Todas_las_sec"))
analizar(2021,"CD1201",c(1:9,"Todas_las_sec"))
analizar(2021,"CD2201",c(33))
analizar(2021,"CD2201",c(1,2,5,6,7,13,14,28))
#para todo adela junto 31,32,33,35,36,37,38,39,40,42,43,44,45,46,47,48,50,"Todas_las_sec"))
analizar(2022,"CD1100",c(1:10,"Todas_las_sec"))
analizar(2022,"CD1201",c(1,3,5,6,7,8,9,10,"Todas_las_sec"))
analizar(2023,"CD1100",c(2,3,4,5,6,7,8,10,"Todas_las_sec"))
analizar(2023,"CD1201",c(2,3,4,5,7,8,9,"Todas_las_sec"))
analizar(2024,"CD1100",c(2,4,6,7,8,"Todas_las_sec"))
analizar(2024,"EH2202",c(1))
file.choose()
