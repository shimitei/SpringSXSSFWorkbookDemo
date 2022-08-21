package com.example.demo;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.util.Strings;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

@Controller
public class PoiController {

	@GetMapping
	public String index() {
		return "index.html";
	}

	@PostMapping("/download")
	public ResponseEntity<StreamingResponseBody> download() throws Exception {

		return ResponseEntity.ok()
				.contentType(MediaType.APPLICATION_OCTET_STREAM)
				.header(HttpHeaders.CONTENT_DISPOSITION,
						"attachment; filename*=UTF-8''" + URLEncoder.encode("エクセル.xlsx", "UTF-8"))
				.body(out -> writeExcel(out));
	}

	private void writeExcel(OutputStream out) throws IOException {
		final List<Map<String, String>> list = getData();
		// template Excel book
		final Path path = new ClassPathResource("/template1.xlsx").getFile().toPath();

		try (final InputStream templateIsReadOnly = Files.newInputStream(path);
				final XSSFWorkbook refBook = (XSSFWorkbook) WorkbookFactory.create(templateIsReadOnly)) {
			// template Sheet
		    final XSSFSheet refSheet = refBook.getSheetAt(1);
		    // template Row(keys)
		    final XSSFRow refKeyRow = refSheet.getRow(0);
		    // template Keys
		    final List<String> keys = new ArrayList<>();
		    for (int i=0; ; i++) {
		    	final XSSFCell cell = refKeyRow.getCell(i);
		    	if (cell == null) break;
		    	keys.add(cell.getStringCellValue());
		    }
		    // template Row(Styles)
		    final XSSFRow refStyleRow = refSheet.getRow(1);
		    // template Styles
		    final List<XSSFCellStyle> styles = new ArrayList<>();
		    for (int i=0; ; i++) {
		    	final XSSFCell cell = refStyleRow.getCell(i);
		    	if (cell == null) break;
				styles.add(cell.getCellStyle());
		    }
		    final int colCount = styles.size();
		    final XSSFCellStyle rowStyle = refStyleRow.getRowStyle();
			// output without template sheet
			refBook.removeSheetAt(1);

			try (final SXSSFWorkbook wb = new SXSSFWorkbook(refBook, -1)) {
				wb.setCompressTempFiles(true);
				// Output Sheet
				final Sheet sh = wb.getSheetAt(0);
				// ref LastRowNum
				int rownum = refBook.getSheetAt(0).getLastRowNum();

				for (final Map<String, String> data : list) {
					rownum++;
					final Row row = sh.createRow(rownum);
					row.setRowStyle(rowStyle);
					row.setHeight(refStyleRow.getHeight());
					// create, style
					for (int cellnum = 0; cellnum < colCount; cellnum++) {
						final Cell cell = row.createCell(cellnum);
				        cell.setCellStyle(styles.get(cellnum));
					}
					// value
					if (data == null) continue;
					for (int cellnum = 0; cellnum < colCount; cellnum++) {
						final String key = keys.get(cellnum);
						if (Strings.isBlank(key)) continue;

						final String val = data.get(key);
						if (Strings.isBlank(val)) continue;
						final Cell cell = row.getCell(cellnum);
						cell.setCellValue(val);
					}
				}

				// output excel book
				wb.write(out);
			}
	    }
	}
	
	private List<Map<String, String>> getData() {
		final String keys[] = {"name1", "name2", "val1", "val2"};
		final String vals[][] = {
				{"にしきぎ", "ちさと", "AB", "9/23"},
				{"いのうえ", "たきな", "A", "8/2"},
				{"なかはら", "みずき", "O", "6/5"},
				{"くるみ", "うぉーるなっと", "AB", "12/16"},
		};

		final List<Map<String, String>> list = new ArrayList<>();
		for (String[] ar : vals) {
			final Map<String, String> map = new HashMap<>();
			map.put(keys[0], ar[0]);
			map.put(keys[1], ar[1]);
			map.put(keys[2], ar[2]);
			map.put(keys[3], ar[3]);
			list.add(map);
		}
		return list;
	}
}
