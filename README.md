"Pedidos" es un conjunto de scripts para procesar la información sobre las compras de diferentes clientes de un liderazgo comercial.

Procesos:

El primero es la descarga de archivos en una carpeta identificada con el nombre CmmAA (C1025 corresponde a la Campaña 10, octubre, de 2025). Cada .xls tendrá el nombre del líder correspondiente en caso de que haya más de uno.

1_xls_xlsx.py buscar en la carpeta asignada y convierte todos los archivos a formato .xlsx necesario para procesos posteriores, SIN borrar los originales.

2_unificar_excels.py genera un archivo de Excel con la información de todos los líderes en un único archivo, que se subirá a la base de datos 'gerencia' de MySQL en dos partes:

	A. 3_subir_clientes.py busca en el archivo los clientes (Nro, Cliente y Líder) y los carga en la base, sin duplicarlos si ya existen.
	B. 4_subir_pedidos.py actualiza los datos de pedidos de esta Campaña.
