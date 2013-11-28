<?php
require_once $path->home_dir ."oes-plugins/PHPExcel.php";
require_once $path->home_dir ."oes-plugins/PHPExcel/Reader/Excel2007.php";
class loadInventoy extends mysql{
	private function uploadXlsFile()
	{
		$this->upldFile = move_uploaded_file($this->tmpLoc,$this->fileLoc);
		if($this->upldFile)
		{
			$this->fileUp = 1;
		}
		else{
			header("Location: " .$this->home_addr ."consadmin/inventario/e1/");
		}
	}
	private function checkNew()
	{
		$this->qry = parent::conexion() ->prepare("SELECT * FROM partes WHERE rawparte=? && cliente=?");
		$this->qry->bind_param('ss',$this->colRes[0],$this->cliente);
		$this->qry->execute();
		$this->qry->store_result();
		$this->n_rows = $this->qry ->num_rows;
		if($this->n_rows >= 1)
		{
			$this->qry->close();
			$this->newRow = 2;
		}
		else	{
			$this->qry->close();
			$this->newRow = 1;
		}
	}
	private function updtRow()
	{
		$this->qrysave = parent::conexion() ->prepare("UPDATE partes SET existencia=?,precio1=?,descripcion=?,parte=? WHERE rawparte=? AND cliente=?");
		$this->qrysave->bind_param('ssssss',$this->colRes[2],$this->colRes[3],$this->colRes[1],$this->colRes[0],$this->colRes[0],$this->cliente);
		$this->qrysave->execute();
		if($this->qrysave->error)
		{
			$this->qrysave->close();
		}
		else{
			$this->qrysave->close();
		}
	}
	private function insertRow()
	{
		$this->qrysave = parent::conexion() ->prepare("INSERT INTO partes (rawparte,parte,existencia,precio1,descripcion,cliente) VALUES(?,?,?,?,?,?)");
		$this->qrysave->bind_param('ssssss',$this->colRes[0],$this->colRes[0],$this->colRes[2],$this->colRes[3],$this->colRes[1],$this->cliente);
		$this->qrysave->execute();
		if($this->qrysave->error)
		{
			$this->qrysave->close();
		}
		else{
			$this->qrysave->close();
		}
	}
	private function defineParams()
	{
		$this->objReader = PHPExcel_IOFactory::createReader($this->excelVersion);
		$this->objReader->setReadDataOnly(true);
		$this->objPHPExcel = $this->objReader->load($this->fileLoc);
		$this->objWorksheet = $this->objPHPExcel->getActiveSheet();
		$this->highestRow = $this->objWorksheet->getHighestRow(); // e.g. 10
		$this->highestColumn = $this->objWorksheet->getHighestColumn(); // e.g 'F'
		$this->highestColumnIndex = PHPExcel_Cell::columnIndexFromString($this->highestColumn); // e.g. 5
	}
	private function updateCustomer()
	{
		$this->currentDate = date("Y-m-d H:i:s");
		$this->qrysave = parent::conexion() ->prepare("UPDATE clientes SET ultact=? WHERE cliente=?");
		$this->qrysave->bind_param('ss',$this->currentDate,$this->cliente);
		$this->qrysave->execute();
		$this->qrysave->close();
	}
	private function getRows()
	{
		$this->row = 2;
		while($this->row <= $this->highestRow) {
			$this->colRes = array();
			$this->col = 0;
			while ( $this->col <= $this->highestColumnIndex) {
				$this->colRes[] =  trim(strip_tags($this->objWorksheet->getCellByColumnAndRow($this->col, $this->row)->getValue()));
				$this->col++;
			}
			$this->checkNew();
			if($this->newRow == 2)
			{
				$this->updtRow();
			}
			elseif($this->newRow == 1)
			{
				$this->insertRow();
			}
			$this->row++;
	      }
	      unlink($this->fileLoc);
	      $this->updateCustomer();
	      header("Location: " .$this->home_addr ."consadmin/inventario/e100/");
	}
	private function defineExt()
	{
		if($this->fileType == "application/vnd.ms-excel")
		{
			$this->excelVersion = "Excel5";
		}
		elseif($this->fileType == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
		{
			$this->excelVersion = "Excel2007";
		}
		else{
			header("Location: " .$this->home_addr ."consadmin/inventario/e2/");
		}
	}
	private function defineElements()
	{
		$this->fileName = $_SESSION['s_idcliente'] ."_" .$_FILES['xlsfile']['name'];
		$this->directory = $this->home_dir ."oes-content/uploads/tmp/";
		if(!is_dir($this->directory))
		{
			mkdir($this->directory,0777);
		}
		$this->fileType = $_FILES['xlsfile']['type'];
		$this->tmpLoc = $_FILES['xlsfile']['tmp_name'];
		$this->fileLoc = $this->directory .$this->fileName;
		$this->cliente = $_SESSION['s_idcliente'];
	}
	public function startLoading()
	{
		$this->defineElements();
		$this->defineExt();
		$this->uploadXlsFile();
		if($this->fileUp == 1)
		{
			$this->defineParams();
			$this->getRows();
		}
	}
}

$upldXls = new loadInventoy();
$upldXls->startLoading();
?>
