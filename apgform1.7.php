<?php 

		
	// Mostrar cuando la operacion de registro es exitosa o no
	$success = "ok.html";
	$error = "no.html";
	
		
	
	// Metodo get o post del formulario
	if($_POST){
		$array = $_POST;
	} else if($_GET){			
		$array = $_GET;
	} else {
			die("You must Access this file through a form.");
	}	

	
	if(!$array['filename']){
		
		$array['filename'] = "form.xls";	//Crear el archivo de excel, que se guarda en la misma carpeta
	
	} else {
		if(!(stristr($array['filename'],".xls"))){
			$array['filename'] = $array['filename'] . ".xls"; //Si el archivo ya existe no lo vuelve a crear
		}
	}
	
	// Defina carácteres de retorn
	$tab = "\t";	//chr(9);
	$cr = "\n";		//chr(13);
	
	if($array){
			
			//Establecer el titulo de las columnas segun los campos del formulario
			$keys = array_keys($array);
			foreach($keys as $key){
				if(strtolower($key) != 'filename' && strtolower($key) != 'title'){ 
					$header .= $key . $tab;
				}
			}
			$header .= $cr;
			
			//Agregar los datos a las columnas
			foreach($keys as $key){
				if(strtolower($key) != 'filename' && strtolower($key) != 'title'){ 

					$array[$key] = str_replace("\n",$lbChar,$array[$key]);
					$array[$key] = preg_replace('/([\r\n])/e',"ord('$1')==10?'':''",$array[$key]);
					$array[$key] = str_replace("\\","",$array[$key]);
					$array[$key] = str_replace($tab, "    ", $array[$key]);
					$data .= $array[$key] . $tab ;
				}
			}
			$data .= $cr;
			
			if (file_exists($array['filename'])) {
				$final_data = $data;		// Si el archivo realmente existe, entonces sólo escriben la información el usuario envió
			} else {
				$final_data = $header . $data;		// Si el archivo no existe , escribir la cabecera (primera línea en Excel con títulos ) en el fichero
			}
			// abrir el archivo y escribir en él
			
			$fp = fopen($array['filename'],"a"); // $fp es ahora el puntero de archivo a archivo $array['filename']
			
			if($fp){
				
				fwrite($fp,$final_data);	//Escribe informacion en el archivo
				fclose($fp);		// Cierra el archivo
				// Exitoso
				header("Location: $success");
			} else {
				// Error
				header("Location: $error");
			}
	}
	
?>
