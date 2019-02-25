package readexport;

import java.awt.Desktop;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.regex.Pattern;

import javax.swing.JOptionPane;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

import com.jxcell.CellFormat;
import com.jxcell.ChartShape;
import com.jxcell.View;

import jxl.Cell;
import jxl.CellView;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.read.biff.File;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

/**
 * 
 * @author 罗蒙and潘超and雍梦婷
 *
 */
public class ExcelHanldle {
	static int[] p = new int[12];
	static String j[] = { "0~9", "10~19", "20~29", "30~39", "40~49", "50~59", "60~69", 
			"70~79", "80~89", "90~99","100" };
	// 在界面输入路径
	static String str = JOptionPane.showInputDialog("请输入文件路径:");

	public static void main(String[] args) throws IOException {
		java.io.File file = new java.io.File(str);

		// 判断输入的路径，是否是文件，是否存在，是否是.xls的Excel文件
		if (!isFiel(str)) {
			JOptionPane.showMessageDialog(null, "输入路径的文件格式不是.xls");
		} else if (!file.isFile()) {
			JOptionPane.showMessageDialog(null, "输入路径不是文件");
		} else if (!file.exists()) {
			JOptionPane.showMessageDialog(null, "输入路径不存在");
		} else {
			// 创建分析表格
			creatExcel();
			JOptionPane.showMessageDialog(null, "已完成" + "D:/test.xls");
			Desktop.getDesktop().open(new java.io.File("D:/test.xls"));
		}
	}

	// 判断文件格式是否为.xls格式的Excel文件
	public static Boolean isFiel(String text) {
		return Pattern.compile("\\S+.+xls").matcher(text).matches();
	}

	private static void creatExcel() {
		// ，存储每个人的成绩
		ArrayList<Integer> list = new ArrayList<>();
		int sum = 0;
		NumberFormat df = NumberFormat.getInstance();
		df.setMaximumFractionDigits(2);
		int count = 0;

		try {
			// 找到需要修改的表
			Workbook wb = Workbook.getWorkbook(new java.io.File(str));
			// 找到表中的第一个
			Sheet read = wb.getSheet(0);
			// 获取表的行数
			int rows = read.getRows();
			// 表的列数
			int clum = read.getColumns();
			// 指定路径
			WritableWorkbook book = Workbook.createWorkbook(new java.io.File("D:/test.xls"));
			// 第几个表被创建
			WritableSheet sheet = book.createSheet("考分频数分布", 0);
		//	jxl.format.CellFormat cf = wb.getSheet(0).getCell(1, 0).getCellFormat();
			WritableCellFormat wc = new WritableCellFormat();
			// 设置居中
			wc.setAlignment(Alignment.CENTRE);
			// 设置边框线
			wc.setBorder(Border.ALL, BorderLineStyle.THIN);
			// 添加新表各列字段（列，行，数据）
			// 第一个表格的框架
			sheet.addCell(new Label(0, 0, "分数段", wc));
			for (int i = 1; i < 12; i++) {
				sheet.addCell(new Label(i, 0, j[i - 1], wc));
			}
			sheet.addCell(new Label(0, 1, "人数", wc));
			sheet.addCell(new Label(0, 2, "比率", wc));

			// 第二个表格的框架
			sheet.addCell(new Label(0, 5, "名and题", wc));
			for (int i = 1; i < clum; i++) {
				sheet.addCell(new Label(i, 5, "题" + i, wc));
			}
			sheet.addCell(new Label(11, 5, "总分", wc));
			String[] a = { "题分值", "得分和", "平均分", "标准差", "难度", "高分平均", "低分平均", "区分度" };
			for (int i = 6; i < 14; i++) {
				sheet.addCell(new Label(0, i, a[i - 6], wc));
			}
			int[] b = new int[clum - 1];
			int[] h = new int[clum];

			// 创建一个新的Excel文件用来存放抽取的数据
			WritableWorkbook newwb = Workbook.createWorkbook(new java.io.File("D:/EXCEL/new.xls"));
			WritableSheet newsheet = newwb.createSheet("数据", 0);
			Sheet newread = newwb.getSheet(0);

			if (rows >= 252) {
				count = (int) ((rows - 2) * 0.2);
				int[] num = new int[rows - 2];
				for (int i = 2; i <= rows - 1; i++) {
					num[i-2] = i;
				}
				// 得出每个人的成绩
				for (int i = rows - 3; i > (int) ((rows - 2) * 0.8); i--) {

					int r = (int) (Math.random() * i) ;
					int t = num[i];
					num[i] = num[r];
					num[r] = t;
					if (i != 1) {
						for (int j = 1; j < clum; j++) {
							// 这个数组存放的是每一行的成绩
							b[j - 1] = (int) ((NumberCell) read.getCell(j, num[i])).getValue();
						}
						// 将抽到的成绩放置新表中
						for (int j = 1; j < clum; j++) {
							Number number = new Number(j, rows - i-1, b[j - 1], wc);
							newsheet.addCell(number);
						}
					}
				}
			} else if (rows < 252 && rows >= 52) {
				count = 50;
				// 得出每个人的成绩
				for (int i = rows - 2; i > rows - 2 - 50; i--) {

					int r = 1 + (int) (Math.random() * i);
					if (i != 1) {
						for (int j = 1; j < clum; j++) {
							// 这个数组存放的是每一行的成绩
							b[j - 1] = (int) ((NumberCell) read.getCell(j, r)).getValue();
						}
						for (int j = 1; j < clum; j++) {
							Number number = new Number(j, rows - i, b[j - 1], wc);
							newsheet.addCell(number);
						}
					}
				}
			} else if (rows < 52) {
				count = rows;
				// 得出每个人的成绩
				for (int i = 2; i < rows; i++) {
					for (int j = 1; j < clum; j++) {
						// 这个数组存放的是每一行的成绩
						b[j - 1] = (int) ((NumberCell) read.getCell(j, i)).getValue();
					}
					for (int j = 1; j < clum; j++) {
						Number number = new Number(j, i, b[j - 1], wc);
						newsheet.addCell(number);

					}
				}
			}

			// 抽取数据的表得行列数
			int newrows = newread.getRows();
			int newclum = newread.getColumns();
            
			//f数组存储每一行的分数
			int[] f = new int[newclum - 1];
			for (int i = 2; i < newrows; i++) {
				for (int j = 1; j < newclum; j++) {
					f[j - 1] = (int) ((NumberCell) newread.getCell(j, i)).getValue();
				}
				for (int w : f) {
					sum += w;
				}
				//list存储每一题的得分数
				list.add(sum);
				Number number = new Number(11, i, sum, wc);
				newsheet.addCell(number);
				sum = 0;
			}
			newclum = newread.getColumns();

			for (int q : list) {
				//p数组存储每个分数段的人数
				if (q < 10) {
					p[0]++;
				} else if (q < 20) {
					p[1]++;
				} else if (q < 30) {
					p[2]++;
				} else if (q < 40) {
					p[3]++;
				} else if (q < 50) {
					p[4]++;
				} else if (q < 60) {
					p[5]++;
				} else if (q < 70) {
					p[6]++;
				} else if (q < 80) {
					p[7]++;
				} else if (q < 90) {
					p[8]++;
				} else if (q < 100) {
					p[9]++;
				} else if (q == 100) {
					p[10]++;
				}
				if (q >= 60) {
					p[11]++;
				}
			}

			for (int i = 1; i < newclum; i++) {
				// 每个分数段的人数
				Number renNumber = new Number(i, 1, p[i - 1], wc);
				sheet.addCell(renNumber);
				// 每个分数段的比率
				DecimalFormat de = new DecimalFormat("0.00%");
				double n = (double) p[i - 1] / (newrows - 2);
				// System.out.println(n);
				Label bilv = new Label(i, 2, de.format(n), wc);
				sheet.addCell(bilv);
			}
			// 得到题分值
			for (int j = 1; j < clum; j++) {
				// 这个数组存放的是一行的题分值
				h[j - 1] = (int) ((NumberCell) read.getCell(j, 1)).getValue();
				h[clum - 1] = 100;
			}

			for (int i = 1; i < newclum; i++) {
				// 题分值添加到试卷分析表中
				Number textNumber = new Number(i, 6, h[i - 1], wc);
				sheet.addCell(textNumber);
			}
			int[] c = new int[newrows - 2];
			double Sum = 0.0;
			double m = 0.0;
			// 列
			for (int i = 1; i < newclum; i++) {
				double s = 0;
				double v = 0.0;
				double d = 0.0;
				double l = 0.0;
				double g = 0.0;
				// 行
				for (int j = 2; j < newrows; j++) {
					// 这个数组存放的是每一列的得分
					c[j - 2] = (int) ((NumberCell) newread.getCell(i, j)).getValue();
				}
				for (int w : c) {
					// 每一题的总得分
					s += w;
					System.out.println(w);
				}
				for (int w : c) {
					// 方差
					v += (w - s / (newrows - 2)) * (w - s / (newrows - 2));

				}
				// 数组内元素进行排序(由小到大)
				Arrays.sort(c);

				// 低分平均分(低分的25%)
				for (int z = 0; z <= (int) (newrows - 2) / 4; z++) {
					l += c[z];
				}
			
				l = Math.round((l / (int) ((newrows - 2) / 4) * 1000)) / 1000.0;
				Number low = new Number(i, 12, l, wc);
				sheet.addCell(low);

				// 高分平均分(高分的25%)
				for (int z = 3 * (newrows - 2) / 4; z < newrows - 2; z++) {
					g += c[z];
				}
				g = Math.round((g / (int) ((newrows - 2) / 4 + 1) * 1000)) / 1000.0;
				Number high = new Number(i, 11, g, wc);
				sheet.addCell(high);

				// 区分度=(高平均分-低平均分)/抽取人数
				double qf = (g - l) / (int) ((newrows - 2) / 4) / h[i - 1];
				qf = Math.round(qf * 1000) / 1000.0;
				Number qfd = new Number(i, 13, qf, wc);
				sheet.addCell(qfd);

				// 每一题的,得分和
				Number numberone = new Number(i, 7, s, wc);
				sheet.addCell(numberone);
				// 平均分
				s = Math.round(s / (newrows - 2) * 100) / 100.0;
				Number averge = new Number(i, 8, s, wc);
				sheet.addCell(averge);
				// 标准差
				d = Math.sqrt(v / (newrows - 2));
				d = Math.round(d * 100) / 100.0;// 保留俩位小数,但数据类型不变
				Number dVar = new Number(i, 9, d, wc);
				sheet.addCell(dVar);
				// 难度=1-平均分/题分
				s = Math.round(s / h[i - 1] * 1000) / 1000.0;
				Number difficult = new Number(i, 10, 1 - s, wc);
				sheet.addCell(difficult);
				// 总结
				String[] last = { "样本数", "信度", "最高分", "最低分", "全距", "及格率" };
				for (int j = 0; j < 6; j++) {
					sheet.addCell(new Label(j, 16, last[j], wc));
				}
				// 样本数
				Number yang = new Number(0, 17, count, wc);
				sheet.addCell(yang);
				// 可信度
				if (i != newclum - 1)
					Sum += v;
				if (i == newclum - 1) {
					m = v;
				}
				v = 0.0;
			}
			// 可信度
			Sum = (newrows - 2) / (newrows - 2 - 1) * (1 - Sum / m);
			Sum = (Math.round(Sum * 100) / 100.0);
			Number number = new Number(1, 17, Sum, wc);
			sheet.addCell(number);
			// 最高分
			Sum = c[newrows - 2 - 1];
			Number high = new Number(2, 17, Sum, wc);
			sheet.addCell(high);
			// 最低分
			m = c[1];
			Number low = new Number(3, 17, m, wc);
			sheet.addCell(low);
			// 全距
			Number juli = new Number(4, 17, Sum - m, wc);
			sheet.addCell(juli);
			// 及格率
			Sum = (double) p[11] / (newrows - 2);
			Sum = (Math.round(Sum * 100) / 100.0);
			Number six = new Number(5, 17, Sum * 100, wc);
			sheet.addCell(six);
			// 创建图片
			creatPNG();

			java.io.File imFile = new java.io.File("D:/EXCEL/test paper.png");
			jxl.write.WritableImage image = new jxl.write.WritableImage(13, 0, 8, 30, imFile);
			// 添加图片
			sheet.addImage(image);

			// 写入
			newwb.write();
			book.write();
			// 关闭
			newwb.close();
			book.close();	

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void creatPNG() throws FileNotFoundException, IOException {
		DefaultCategoryDataset dataset = new DefaultCategoryDataset();

		// 添加数据
		for (int i = 0; i < 11; i++) {
			dataset.addValue(p[i], "grade", j[i]);
		}
		JFreeChart chart = ChartFactory.createBarChart("test paper", "Grade", // X轴名称
				"peopleNumber", // Y轴名称
				dataset, // 数据
				PlotOrientation.VERTICAL, // 图标方向：垂直
				true, // 是否生成图例
				false, // 是否生成工具
				false); // 是否产生超链接
		FileOutputStream fos = null;

		try {
			// 输出路径
			fos = new FileOutputStream("D:/EXCEL/test paper.png");
			ChartUtilities.writeChartAsPNG(fos, chart, 800, 1000);

		} finally {
			// 关闭,清缓存
			fos.close();
		}
	}

}
