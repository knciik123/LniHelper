package lniHelper;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;
import java.util.Spliterator;
import java.util.Spliterators;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Index {
	public static int cpu = Runtime.getRuntime().availableProcessors();
	public static ThreadPoolExecutor executor = (ThreadPoolExecutor) Executors.newFixedThreadPool(cpu);

	@SuppressWarnings("resource")
	public static void main(String[] args) throws Exception {
		if (args == null || args.length<1) 
			System.out.println("��������,��鿴readme.md");
		
		String basePath = args[0].replace("\"", "");
		String inputPath = basePath + "\\excel\\excel.xlsx";
		String outputPath =basePath + "\\table";
		String logPath = basePath + "\\excel\\log.txt";
		
		Optional.ofNullable(inputPath).map(File::new).map(excelFile -> {
			Hashtable<String, String> namePath = new Hashtable<>();
			namePath.put("ability.ini", outputPath + "\\ability.ini");
			namePath.put("buff.ini", outputPath + "\\buff.ini");
			namePath.put("destructable.ini", outputPath + "\\destructable.ini");
			namePath.put("doodad.ini", outputPath + "\\doodad.ini");
			namePath.put("item.ini", outputPath + "\\item.ini");
			namePath.put("unit.ini", outputPath + "\\unit.ini");
			namePath.put("misc.ini", outputPath + "\\misc.ini");
			namePath.put("txt.ini", outputPath + "\\txt.ini");
			namePath.put("upgrade.ini", outputPath + "\\upgrade.ini.ini");
			
			File logFile = new File(logPath);

			if (!logFile.getParentFile().exists())
				logFile.getParentFile().mkdir();

			if (!logFile.exists()) 
				try {
					logFile.createNewFile();
				} catch (IOException e) {
					e.printStackTrace();
				}

			if (excelFile.isFile() && excelFile.exists()) 
				try {
					Workbook wb = new XSSFWorkbook(excelFile);
					Iterator<Sheet> car = wb.sheetIterator();

					while (car.hasNext()) {
						Sheet sheet = car.next();
						String name = sheet.getSheetName();
						executor.execute(() -> {
							try {
								parseSheet(name, namePath.get(name), sheet);
							} catch (RuntimeException e1) {
								System.out.println("sheet�����쳣  " + name + " " + e1.getClass());
							} catch (Exception e2) {
								System.out.println("sheetδ����  " + name + " " + e2.getMessage());
							}
						});
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			else
				return null; //����null����orElseThrow
			
			return excelFile;
		}).orElseThrow(()->new FileNotFoundException("δ�ҵ�excel.xlsx"));
		executor.shutdown();
	}

	public static void parseSheet(String name, String path, Sheet sheet) throws Exception {
		Optional.ofNullable(path).filter(Utils::isNotEmpty).orElseThrow(() -> new Exception("sheet��������"));

		StringBuffer buffer = new StringBuffer();
		List<String> titles = new ArrayList<>();

		// ��ȡ�ֶ���titles
		Optional.ofNullable(sheet.getRow(1)).map(row -> {
			StreamSupport.stream(Spliterators.spliteratorUnknownSize(row.iterator(), Spliterator.ORDERED), false)
					.forEach(cell -> {
						Optional.ofNullable(cell).map(Utils::cellString).ifPresent(titles::add);
					});
			return row;
		}).orElseThrow(() -> new Exception("�յı��"));
		//titles.stream().forEach(System.out::println); //��ӡtitles
		
		// ����ÿһ��
		try {
			StreamSupport.stream(Spliterators.spliteratorUnknownSize(sheet.iterator(), Spliterator.ORDERED), false)
				.skip(2).forEach(row -> {
					Optional.of(row.getCell(1)).map(Utils::cellString).filter(Utils::isNotEmpty).ifPresent(v->{
						buffer.append(String.format("[%s]\n", v)); //ƴ��ID
					});
					
					StreamSupport.stream(Spliterators.spliteratorUnknownSize(row.iterator(), Spliterator.ORDERED),false)
						.skip(2).forEach(cell->{
							try {
								Optional.ofNullable(cell).map(Utils::getValue).filter(Utils::isNotEmpty).ifPresent(v->{
									buffer.append(String.format("%s = %s\n", titles.get(cell.getColumnIndex()), v)); //ƴ������
								});
							} catch (Exception e) {
								
							}
						});
					buffer.append("\n");
				});

		} catch (Exception e) {

		}

		// ����ļ�
		output(new File(path), buffer.toString());
	}

	public static void output(File file, String str) throws Exception {
		if (!file.getParentFile().exists())
			file.getParentFile().mkdirs();

		if (file.exists())
			file.delete();

		file.createNewFile();

		try (FileOutputStream fos = new FileOutputStream(file);
				OutputStreamWriter osw = new OutputStreamWriter(fos, "UTF-8");
				BufferedWriter bw = new BufferedWriter(osw);) {
			bw.write(str);
		} catch (Exception e) {
			throw new Exception("д��ʧ�� " + file.getPath());
		}
	}
}

class Utils {
	private static final DecimalFormat DECIMAL_FORMAT = new DecimalFormat("0");// ��ʽ�� numberΪ��
	private static final DecimalFormat DECIMAL_FORMAT_PERCENT = new DecimalFormat("##.00%");// ��ʽ���ֱȸ�ʽ�����治��2λ����0����
	private static final DecimalFormat DECIMAL_FORMAT_NUMBER = new DecimalFormat("0.00E000"); // ��ʽ����ѧ������
	private static final Pattern POINTS_PATTERN = Pattern.compile("0.0+_*[^/s]+"); // С��ƥ��
	private static final ThreadLocal<DateFormat> FAST_DATE_FORMAT = new ThreadLocal<DateFormat>() {
		@Override
		public SimpleDateFormat initialValue() {
			return new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		}
	};

	public static Row getRow0(Sheet s) {
		return s.getRow(0);
	}

	public static Cell getCell0(Row w) {
		return w.getCell(0);
	}

	public static Row getRow1(Sheet s) {
		return s.getRow(1);
	}

	public static Cell getCell1(Row w) {
		return w.getCell(1);
	}

	public static String getValue(Cell cell) {
		Object value = null;
		switch (cell.getCellTypeEnum()) {
		case _NONE:
			break;

		case STRING:
			value = cell.getStringCellValue();
			if ("#".equals(value))
				value = "";
			else if ("*".equals(value))
				value = "\"\"";
			else if ("$".equals(value))
				value = null;
			else
				value = "\"" + cell.getStringCellValue() + "\"";
			break;

		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell))// ����
				value = FAST_DATE_FORMAT.get().format(DateUtil.getJavaDate(cell.getNumericCellValue()));
			else if ("@".equals(cell.getCellStyle().getDataFormatString())
					|| "General".equals(cell.getCellStyle().getDataFormatString())
					|| "0_ ".equals(cell.getCellStyle().getDataFormatString()))
				value = DECIMAL_FORMAT.format(cell.getNumericCellValue());// �ı� or ���� or ������ֵ
			else
				value = cell.getNumericCellValue(); // ֱ����ʾ
			break;
		case BOOLEAN:
			value = cell.getBooleanCellValue();
			break;
		case BLANK:
			value = "\"\"";
			break;
		default:
			value = cell.toString();
		}
		return value.toString();
	}

	public static boolean isEmpty(String s) {
		return s == null || "".equals(s);
	}

	public static boolean isNotEmpty(String s) {
		return !isEmpty(s);
	}

	public static String cellString(Cell c) {
		return c.toString();
	}
}