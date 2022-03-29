# Generador-de-Platilla-Amazon-c++-
Utiliza la librería OpenXLSX para obtener y acomodar información de un excel a otro.

Para que el programa funcione tenemos que tener una base de datos en excel que es basicamente una tabla de donde se estaran tomando todos los datos necesarios para generar un formato tal cual lo pide amazon para subir productos, en este caso se pretende generar plantillas para subir cuadros decorativos a Amazon.

Aqui puedes obtener mas información sobre la libreria OpenXLSX  https://github.com/troldal/OpenXLSX

El programa fue creado en Visual Studio 2019 creando un proyecto Windows Forms C++ CRL.

El resultado final es el siguiente. 

![image](https://user-images.githubusercontent.com/22418357/160686782-e220c908-8a60-460d-b594-2941c876c21c.png)

En la carpeta del proyecto tuve que agregar la libreria OpenXLS para que funcionara.

![image](https://user-images.githubusercontent.com/22418357/160687707-8a836251-0e78-4845-bd5e-7e7f9d17fab0.png)

En la cartpeta debug donde resulta el exe debemos compiar la dll OpenXLSXD.dll para que pueda compilar.

![image](https://user-images.githubusercontent.com/22418357/160688321-fce7d514-34ff-4d9a-ae29-0807c7aa6604.png)

Donde se va a ejecutar el programa estoy colocando la base de datos DCANVAS(alan).xlsx de donde saco toda la información , para despues procesarla y generar otro libro de excel con el formato que  nos pide Amazon en su platilla para subir productos.
