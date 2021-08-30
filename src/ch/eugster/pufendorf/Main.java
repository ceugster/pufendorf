package ch.eugster.pufendorf;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main 
{
	public static void main(String[] args) throws InvalidFormatException, IOException
	{
		Main main = new Main();
		File[] files = main.resolve(args);
		for (int i = 0; i < files.length; i++)
		{
			System.out.println("Starting with workbook " + files[i].getName() + " (" + (i + 1) + ")...");
			main.loopSheets(files[i]);
			System.out.println("Ending with workbook " + files[i].getName() + " (" + (i + 1) + ")...");
		}
	}

	public File[] resolve(String[] args)
	{
		List<File> files = new ArrayList<File>(); 
		if (args.length == 0)
		{
			File directory = new File(System.getProperty("user.dir"));
			files = Arrays.asList(directory.listFiles(new FilenameFilter()
			{
				@Override
				public boolean accept(File dir, String name) 
				{
					return name.endsWith(".xlsx") && !name.toLowerCase().contains("matrix");
				}
			}));
			System.out.println("Valid files found " + files.size() + " out of " + (args == null ? 0 : files.size()));
		}
		else
		{
			for (int i = 0; i < args.length ; i++)
			{
				File file = new File(args[i]);
				if (file.exists() && file.getName().endsWith(".xslx") && !file.getName().toLowerCase().contains("matrix"))
				{
					files.add(file);
				}
			}
			System.out.println("Valid files found " + files.size() + " out of " + (args == null ? 0 : args.length));
		}
		return files.toArray(new File[0]);
	}

	private void loopSheets(File file) throws InvalidFormatException, IOException
	{
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		for (int i = 0 ; i < workbook.getNumberOfSheets(); i++)
		{
			Sheet sheet = workbook.getSheetAt(i);
			if (sheet.getSheetName().toLowerCase().contains("-normalized") || sheet.getSheetName().toLowerCase().contains("-processed"))
			{
				workbook.removeSheetAt(i);
				i = 0;
			}
		}
		int numberOfSheets = workbook.getNumberOfSheets();
		for (int i = 0 ; i < numberOfSheets ; i++)
		{
			System.out.println("   Processing sheet " + workbook.getSheetAt(i).getSheetName() + " (" + (i + 1) + ")...");
			XSSFSheet normalizedSheet = normalizeCountries(workbook.getSheetAt(i));
			XSSFSheet targetSheet = workbook.createSheet(workbook.getSheetName(i) + "-processed");
			this.generateMatrix(normalizedSheet, targetSheet);
		}
		File targetFile = new File(file.getParent() + File.separator + file.getName().replace(".xlsx", "-matrix.xlsx"));
		if (targetFile.exists() && targetFile.isFile())
		{
			targetFile.delete();
		}
		OutputStream os = new FileOutputStream(file.getAbsoluteFile());
		workbook.write(os);
		workbook.close();
		os.close();
	}
	
	private XSSFSheet normalizeCountries(XSSFSheet source) throws IOException
	{
		XSSFSheet target = source;
		if (source.getRow(source.getFirstRowNum()).getCell(0) != null && !source.getRow(source.getFirstRowNum()).getCell(0).getRawValue().trim().isEmpty())
		{
			// Create new target sheet
			target = source.getWorkbook().createSheet(source.getSheetName() + "-normalized");
			// Create first row in target sheet
			Row targetRow = target.createRow(0);
			// Loop over source rows
			Iterator<Row> rowIterator = source.rowIterator();
			while (rowIterator.hasNext())
			{
				// Goto next row
				Row sourceRow = rowIterator.next();
				if (sourceRow.getRowNum() == source.getFirstRowNum())
				{
					// Copy the items of the source's first row to target 0 row
					Iterator<Cell> cellIterator = sourceRow.cellIterator();
					while (cellIterator.hasNext())
					{
						Cell sourceCell = cellIterator.next();
						Cell targetCell = targetRow.createCell(sourceCell.getColumnIndex() + 1);
						targetCell.setCellValue(sourceCell.getStringCellValue());
					}
				}
				else
				{
					// Loop over cells in current row
					Iterator<Cell> cellIterator = sourceRow.cellIterator();
					while (cellIterator.hasNext())
					{
						Cell sourceCell = cellIterator.next();
						if (sourceCell.getCellType().equals(CellType.STRING))
						{
							findOrUpdateCountry(target, sourceCell);
						}
					}
				}
			}
			source.getWorkbook().write(source.getWorkbook().getPackagePart().getOutputStream());
		}
		return target;
	}
	
	private void findOrUpdateCountry(XSSFSheet target, Cell sourceCell)
	{
		// Look for existing country name in first column
		Iterator<Row> rowIterator = target.rowIterator();
		boolean found = false;
		// Skip first row because it is the header row
		if (rowIterator.hasNext())
		{
			rowIterator.next();
		}
		// From now on there are data rows
		if (rowIterator.hasNext())
		{
			while (rowIterator.hasNext())
			{
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				if (cellIterator.hasNext())
				{
					Cell countryCell = cellIterator.next();
					if (countryCell.getCellType().equals(CellType.STRING))
					{
						if (countryCell.getStringCellValue().equals(sourceCell.getStringCellValue())) 
						{
							Cell targetCell = row.getCell(sourceCell.getColumnIndex() + 1);
							if (targetCell == null)
							{
								targetCell = row.createCell(sourceCell.getColumnIndex() + 1);
							}
							targetCell.setCellValue(1d);
							found = true;
						}
					}
				}
			}
		}
		if (!found)
		{
			Cell countryCell = target.createRow(target.getLastRowNum() + 1).createCell(0);
			countryCell.setCellValue(sourceCell.getStringCellValue());
			Cell targetCell = countryCell.getRow().getCell(sourceCell.getColumnIndex() + 1);
			if (targetCell == null)
			{
				targetCell = countryCell.getRow().createCell(sourceCell.getColumnIndex() + 1);
				targetCell.setCellValue(1d);
			}
		}
	}
	
	private void generateMatrix(XSSFSheet source, XSSFSheet target)
	{
		createTargetHeaders(source, target);
		addConnections(source, target);
	}
	
	private void createTargetHeaders(XSSFSheet source, XSSFSheet target)
	{
		// Write countries to first row in target and to first column from row 2 on
		Row headerRow = null;
		Iterator<Row> rowIterator = source.rowIterator();
		if (rowIterator.hasNext())
		{
			headerRow = rowIterator.next();
		}
		int i = 1;
		while (rowIterator.hasNext())
		{
			Row sourceRow = rowIterator.next();
			Iterator<Cell> cellIterator = sourceRow.cellIterator();
			if (cellIterator.hasNext())
			{
				Row targetRow = target.getRow(headerRow.getRowNum());
				if (targetRow == null)
				{
					targetRow = target.createRow(headerRow.getRowNum());
				}
				// Set Row Header Cell
				Cell sourceCell = cellIterator.next();
				Cell targetCell = targetRow.createCell(i++);
				targetCell.setCellValue(sourceCell.getStringCellValue());
				// Set country to first column of current row
				target.createRow(sourceRow.getRowNum()).createCell(0).setCellValue(sourceCell.getStringCellValue());
			}
		}
	}
	
	private void addConnections(XSSFSheet source, XSSFSheet target)
	{
		
		for (short i = 1; i < source.getRow(0).getLastCellNum() + 1 ; i++)
		{
			Iterator<Row> rowIterator = source.rowIterator();
			if (rowIterator.hasNext())
			{
				// This is the header row
				rowIterator.next();
			}
			List<Integer> rows = new ArrayList<Integer>();
			while (rowIterator.hasNext())
			{
				Row row = rowIterator.next();
				Cell cell = row.getCell(i);
				if (cell != null)
				{
					if (cell.getCellType().equals(CellType.NUMERIC) && cell.getNumericCellValue() == 1d)
					{
						rows.add(Integer.valueOf(row.getCell(i).getRowIndex()));
					}
				}
			}
			if (rows.size() > 1)
			{
				int start = 0;
				for (int f = start; f < rows.size(); f++)
				{
					for (int s = start + 1; s < rows.size(); s++)
					{
						Cell cell = target.getRow(rows.get(f)).getCell(rows.get(s));
						if (cell == null)
						{
							cell = target.getRow(rows.get(f)).createCell(rows.get(s));
							cell.setCellValue(1d);
						}
						else if (cell.getCellType().equals(CellType.NUMERIC))
						{
							cell.setCellValue(cell.getNumericCellValue() + 1d);
						}
						else
						{
							cell.setCellValue(1d);
						}
					}
					start++;
				}
			}
		}
	}
}
