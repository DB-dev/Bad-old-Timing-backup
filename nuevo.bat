@ECHO OFF

set debug=1

rem TODO
rem CONTEMPLAR LOS CASOS DE CAMBIO DE MES Y DE AÑO

set plantilla=0_plantilla.xlsx
set archivo=tiempos_hoy.xlsx
set archivo_debug=tiempos_hoy_debug.xlsx

set ymd=%date:~-4%-%date:~3,2%-%date:~0,2%
set ym=%date:~-4%-%date:~3,2%
set /a dia_ayer=%date:~0,2% - 1
set /a mes_pasado=%date:~3,2% - 1
set /a ano_pasado=%date:~-4% - 1
set ayer=%ym%-%dia_ayer%

set ymd_d=2017-03-01
set ym_d=2017-03
set y_d=2017
set /a dia_ayer_d=0
set /a mes_pasado_d=2
set /a ano_pasado_d=2016
set ayer_d=2017-03-00

set y_final=0
set m_final=0
set d_final=0

SET mypath=%~dp0

Rem Si el debug no está activo...
IF NOT %debug%==1 (
	rem Existe el archivo del escritorio de este usuario?
	cd %appdata%
	cd ../../Desktop

	rem Si existe lo movemos, cambiamos el nombre y lo guardamos en su carpeta
	IF EXIST %archivo% ( 
		move /-Y %archivo% %mypath%
		cd %mypath%
		ren %archivo% "%ayer%.xlsx"
		
		rem Comprobamos si existe ya la carpeta y lo movemos, sino, pues creamos la carpeta
		IF EXIST %ym% (
			move /-Y "%ayer%.xlsx" %ym%
		) ELSE (
			mkdir %ym%
			move /-Y "%ayer%.xlsx" %ym%
		)	
	)

	cd %mypath%
	rem Creamos el nuevo archivo de hoy y lo llevamos al escritorio
	copy %plantilla% %archivo%

	cd %appdata%
	cd ../../Desktop

	move %mypath%\%archivo% .\
) 

rem Si el debug está activo
else if %debug%==1 (
	rem Existe el archivo del escritorio de este usuario?
	cd %appdata%
	cd ../../Desktop

	rem Si existe lo movemos, cambiamos el nombre y lo guardamos en su carpeta
	IF EXIST %archivo_debug% ( 
		move /-Y %archivo_debug% %mypath%
		cd %mypath%
		
		rem Comprobamos cambio de mes
		if %dia_ayer_d%==0 (
			
			rem Comprobamos cambio de año (si el mes era enero, el mes pasado es 0 por lo que lo pasamos al 12 y el año al anterior)
			if %mes_pasado_d%==0 (
				set y_final=%ano_pasado_d%
				set m_final=12
				rem falta averiguar si el ultimo archivo los dos ultimos caracteres fueron 29 o 30
			) else (
				rem set archivo_final=%y_d%-%mes_pasado_d%-%ayer_d%
				set y_final=%y_d%
				set m_final=%mes_pasado_d%
				rem falta averiguar si el ultimo archivo los dos ultimos caracteres fueron 29 o 30
			)
			set archivo_final=y_final-m_final-d_final
		)
		
		ren %archivo_debug% "%ayer_d%.xlsx"
		
		rem Comprobamos si existe ya la carpeta y lo movemos, sino, pues creamos la carpeta
		IF EXIST %ym_d% (
			move /-Y "%ayer_d%.xlsx" %ym_d%
		) ELSE (
			mkdir %ym_d%
			move /-Y "%ayer_d%.xlsx" %ym_d%
		)	
	)

	cd %mypath%
	rem Creamos el nuevo archivo de hoy y lo llevamos al escritorio
	copy %plantilla% %archivo_debug%

	cd %appdata%
	cd ../../Desktop

	move %mypath%\%archivo_debug% .\
	
	rem volvemos a la carpeta por motivos de debug
	cd %mypath%
)