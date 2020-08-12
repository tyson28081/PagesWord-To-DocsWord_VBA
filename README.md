# PagesWord-To-DocsWord_VBA
With this ***VBA MACRO WORD*** Resource, we are going to convert each page of a Word to independent .DOCX documents. For example, if I have a Word Document with 7 pages, these 7 pages will be extracted separately as new Word Documents; I mean 7 Word documents apart from the original.

tyson_77newday@hotmail.com

+51 970967059    *whatsapp*

**Ritter77** :virgo:    *I speack english too*

Other Projects:
	http://www.colinaazul.tk/

If you can *donate* something to me, ***I would really appreciate it***, life has its bad sides. </br>
<a href="https://www.paypal.com/paypalme/PeterHanz77" title="Donations Paypal" target="_blank"><img src="PAYPAL Credit Card.jpg" width="200" height="134"></a> </br>
<a href="https://www.paypal.com/paypalme/PeterHanz77" title="Donations Paypal" target="_blank">paypal.me/PeterHanz77</a>
[paypal.me/PeterHanz77](https://www.paypal.com/paypalme/PeterHanz77){target="_blank"} </br>
[Go to this page](https://www.paypal.com/paypalme/PeterHanz77/?target=_blank) </br>
[paypal.me/PeterHanz77](https://www.paypal.com/paypalme/PeterHanz77){:target="_blank"} </br>
[Link](https://www.paypal.com/paypalme/PeterHanz77/target="_blank")
---
---

***HOW TO APPLY?*** *in Spanish*

1. Abrir el Documento *PW TIENDA de CAMARAS FOTOGRAFICAS.docx* </br>

2. fichero: **Archivo** > opciones > personalizar cinta de opciones > </br>
	- check: Programador (panel derecho) </br>
	
3. fichero: **Programador** > visual basic > panel izquierdo > *Seleccionar* </br>
	- Project *PW TIENDA de CAMARAS FOTOGRAFICAS* </br>
		- Microsoft Word Objetos </br>
			- this document </br>
	- Pegar aqui el código del archivo </br>
		- *MACRO PARA DIVIDIR PAGINAS A DOCUMENTOS WORD.txt* </br>
		
4. Modificar las siguientes lineas: </br>
	- **Documents("PW TIENDA de CAMARAS FOTOGRAFICAS.docm").Activate** </br>
	- **"D:\..." </br>**
	
	- Bien te explico en la primera linea *Activamos* el documento que vamos a usar para dividir sus paginas; tiene que ser el mismo nombre que usaras 
		para guardar la macro. </br>
	- Y la segunda linea es la ruta donde se guardaran los nuevos documentos words extraidos del documento original. </br>
	
5. *GUARDAR DOCUMENTO HABILITADO PARA MACROS* </br>
	- Ahora **CTRL + S** para guardar </br>
	- Saldra un mensaje indicando que el documento tiene caracteristicas *Macros*, Dale ***NO*** para *guardar con macros*
	- Fichero: **Archivo** > guardar como > examinar </br>
		- Buscar la carpeta para guardar </br>
		- Nombre --- 		***PW TIENDA de CAMARAS FOTOGRAFICAS*** </br>
		- Tipo de archivo --- 	***documento habilitado con macros de Word (*.DOCM)*** </br>
		- Guardar </br>
		
6. Cerrar el *Visual Basic* y el documento *PW TIENDA de CAMARAS FOTOGRAFICAS.docx*, este sin guardar cambios. </br>

7. Abrimos ahora el nuevo documento *Macros* ***PW TIENDA de CAMARAS FOTOGRAFICAS.docm***. </br>

8. ***Visual Basic*** </br>
	- Ir a la consola de Visual Basic donde pegamos el codigo funcion ***Sub SAVE_INDEPENDENT_SHEETS()*** </br>
	- Presionar ***Ejecutar*** (icono de play) o ***F5***. </br>
	- Los nuevos documentos se guardaran en la *carpeta destino* que elegistes para guardar. </br>
	
9. Listo es todo... ***MUCHAS GRACIAS!!***



