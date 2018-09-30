package zsyw;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.text.NumberFormat;
import java.util.ArrayList;

import javax.print.DocFlavor.STRING;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JTextField;

public class zsyw extends JFrame {
	private static int ResultChuHuo = 0;
	private static int ResultTuiHuo = 0;
	private static String ResultPercent = "";

	 private static ArrayList<kaoqingInput> listInput = null;
	//private static ArrayList<Tuihuo> listTuihuo = null;

 

	private JPanel jp1;
	private JPanel jp2;
	private JPanel jp3;
	private JPanel jp4;
	private JPanel jp5;
	private JPanel jp6;
	private JPanel jp7;
	
	private JTextArea jta1;
	private JTextArea jta2;
	private JComboBox jcb;
	private JTextField jtf1;
	private JTextField jtf2;
	private JTextField jtf3;
	private JButton jb1;
	private JButton jb2;
	private JComboBox jcbb;  //������  

	private JLabel jlb1, jlb2, jlb3, jlb4;
	

	// int chuhuoCount = 0;
	// int tuihuoCount = 0;
	public static void main(String[] args) {
		// TODO Auto-generated method stub

		System.out.println("test!");
		String FilePath = "D:/yewu.xls";
		File f = new File(FilePath);// Դ�ļ�

		zsyw yw = new zsyw();
		yw.showWindow();

	}

	public void showWindow() {
		this.setTitle("我是薄荷啊呀呀 Beta 1.1");// ��������Ϊ�����족
		this.setSize(600, 300);// �����СΪ600*500
		this.setLocation(200, 200);// ����λ�þ���Ļ���Ͻ�Ϊ200*200
		this.setDefaultCloseOperation(EXIT_ON_CLOSE);// �رմ�������
		this.setResizable(true);// �����СΪ���ɱ䣬falseΪ���ɱ䣬trueΪ�ɱ�

		jlb1 = new JLabel("�ͻ�");
		jlb2 = new JLabel("�ͺ�");
		jlb3 = new JLabel("�쳣");

		jp1 = new JPanel();// ��������
		jp2 = new JPanel();// ��������
		jp3 = new JPanel();// ��������
		jp4 = new JPanel();// ��������
		jp5 = new JPanel();// ��������
		jp6 = new JPanel();// ��������
		jp7 = new JPanel();// ��������
		
		// jcb=new JComboBox(new String[]{"����1","����2","����3","����4"});//���ö�ѡ��
		jtf1 = new JTextField(40);// ������Ϣ�༭����ʾ����Ϊ20����
		jtf2 = new JTextField(40);// ������Ϣ�༭����ʾ����Ϊ20����
		//jtf3 = new JTextField(50);// ������Ϣ�༭����ʾ����Ϊ20����
		//jcbb = new JComboBox(StringsYiChang);
		jb1 = new JButton("Test");// ������ť
		jb2 = new JButton("Open");// ������ť
		//this.setLayout(new GridLayout(6, 1));
		this.setLayout(new FlowLayout());
		// jp1.add(jcb);
		jp1.add(jtf1);
		jp1.add(jb2);
		

		jp2.add(jlb2);
		jp2.add(jtf2);

		jp3.add(jlb3);
		//jp3.add(jcbb);
		//jp3.add(jtf3);

		jta1 = new JTextArea();// �����ı���
		Font font=new Font("����",Font.PLAIN,18);
				
		jta1.setFont(font);
		
		jta1.setSize(getMaximumSize());
		jta1.setText("");
		
		jta2 = new JTextArea();// �����ı���
		jta2.setFont(font);
		//jta1.setSize(getMaximumSize());
		//jta2.setBackground(Color.red);
		
		jp4.add(jb1);
		
		jp5.add(jta1);
		jp6.add(jta2);
		jb2.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JFileChooser jfc=new JFileChooser();
				jfc.setFileSelectionMode(JFileChooser.FILES_ONLY );
				jfc.showDialog(new JLabel(), "选择");
				File file=jfc.getSelectedFile();
				if(file.isDirectory()){
					System.out.println("文件夹:"+file.getAbsolutePath());
				}else if(file.isFile()){
					System.out.println("文件:"+file.getAbsolutePath());
				}
				System.out.println(jfc.getSelectedFile().getName());

				jtf1.setText(file.getAbsolutePath());
			}
		});
		
		
		jb1.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				System.out.println("addActionListener!");

				
				String filePath = jtf1.getText();
				if(filePath.isEmpty()) {
					jta1.setText("先选择文件");
					return;
				}
				
				
				//ResultChuHuo = 0;
				//ResultTuiHuo = 0;

//				
//				 if (listChuhuo == null) {
						try {
							ReadDataToList(filePath);
//							
							WriteDateFromList();
							jta1.setText("操作完成\nD:\\1.xls");
//							FilterData();
//
//							ShowResult(); 
						} catch (Exception e1) {
//							// TODO Auto-generated catch block
//							listChuhuo = null;
							jta1.setText("出现错误了");
							System.out.println(e1.getMessage());
							jta1.setBackground(Color.red);
//							e1.printStackTrace();
//							
//							
						} 
//					}else {
//						FilterData();
//
//						ShowResult(); 
//					}
//				 
//				 if (listChuhuo == null) {
//					 jta1.setText("��ȡ�ļ�ʧ��, ��ȷ���ļ�����\nD:\\\\ͳ��.xls");
//					}else {
//						
//					}
//					
				

				
				// jta1.setText("����: " + "\n�˻�: " + "\n�˻���: " ) ;

			}

		 

		});


		this.add(jp1);
		//this.add(jp2);
		//this.add(jp3);
		this.add(jp4);
		this.add(jp5);
		//this.add(jp6);
		this.setVisible(true);
	}

	protected void WriteDateFromList() throws IOException {
		// TODO Auto-generated method stub
		WriteExcel.writeExcel(listInput, 4, "D:\\1.xls");
		
	}

	protected void ShowResult() {
		System.out.println("ShowResult!");

		//float percent = (float) ResultTuiHuo / (float) ResultChuHuo;
		
//		BigDecimal d1 = new BigDecimal(Integer.toString(ResultTuiHuo));
//        BigDecimal d2 = new BigDecimal(Integer.toString(ResultChuHuo));
//        BigDecimal d100 = new BigDecimal(Integer.toString(100));
//        
//        ResultPercent = (d1.multiply(d100).divide(d2,6,BigDecimal.ROUND_HALF_UP)).toString();
		//NumberFormat nf =NumberFormat.getInstance();

		//nf.setGroupingUsed(false);

		//ResultPercent= nf.format(Float.toString(percent));
		
		// TODO Auto-generated method stub
		//jta1.setText("�ͻ�: " + jtf1.getText() +
//				"\n�ͺ�: " + jtf2.getText() +
//				"\n����: " + ResultChuHuo + 
//				"\n�˻�: " + ResultTuiHuo + 
//				"\n�쳣����: " + jcbb.getSelectedItem() +
//				 "\n�˻���: "+ ResultPercent + "%");
//		 
//		jta1.setBackground(Color.cyan);

	}

//	protected void FilterData() {
//		System.out.println("FilterData!");
//		// TODO Auto-generated method stub
//
//		String costemName = jtf1.getText();
//		boolean hasNotCostem = jtf1.getText().isEmpty() || jtf1.getText().toString().equals("");
//
//		String typeName = jtf2.getText();
//		boolean hasNotType = jtf2.getText().isEmpty() || jtf2.getText().toString().equals("");
 
//		
//		String xianxiang = (String) jcbb.getSelectedItem();//jtf3.getText();
//		boolean hasNotXianxiang = xianxiang.equals("����"); //jtf3.getText().isEmpty() || jtf3.getText().toString().equals("");
//		
//		
//		
//		// 3����,��������
//		if (hasNotCostem == true && hasNotType == true && hasNotXianxiang == true) {
//			for (Chuhuo item1 : listChuhuo) {
//				ResultChuHuo += item1.getCount();
//
//			}
//
//			for (Tuihuo item2 : listTuihuo) {
//				ResultTuiHuo += item2.getTotalCount();
//			}
//
//			return;
//		}
//
//		// ��ʼ����
//		// һ�ι��˿ͻ�
//		if (hasNotCostem == false) {
//			for (Chuhuo item : listChuhuo) {
//				if (costemName.equals(item.getCostem())) {
//					listChuhuoFilter1.add(item);
//				}
//			}
//
//			for (Tuihuo item : listTuihuo) {
//				if (costemName.equals(item.getCostem())) {
//					listTuihuoFilter1.add(item);
//				}
//			}
//
//		} else {
//			listChuhuoFilter1 = listChuhuo;
//			listTuihuoFilter1 = listTuihuo;
//		}
//
//		System.out.println("һ�ι���!" + " listChuhuoFilter1.size()=" + listChuhuoFilter1.size());
//		System.out.println("һ�ι���!" + " listTuihuoFilter1.size()=" + listTuihuoFilter1.size());
//
//		// ���ι����ͺ�
//		if (hasNotType == false) {
//			for (Chuhuo item : listChuhuoFilter1) {
//				if (typeName.equals(item.getType())) {
//					listChuhuoFilter2.add(item);
//				}
//			}
//
//			for (Tuihuo item : listTuihuoFilter1) {
//				if (costemName.equals(item.getCostem())) {
//					listTuihuoFilter2.add(item);
//				}
//			}
//		} else {
//			listChuhuoFilter2 = listChuhuoFilter1;
//			listTuihuoFilter2 = listTuihuoFilter1;
//		}
//
//		for (Chuhuo item : listChuhuoFilter2) {
//			ResultChuHuo += item.getCount();
//		}
//		
//		for (Tuihuo item : listTuihuoFilter2) {
//			ResultTuiHuo += item.getTotalCount();
//		}
//		
//		System.out.println("���ι���!" + " listChuhuoFilter2.size()=" + listChuhuoFilter2.size());
//		System.out.println("���ι���!" + " listTuihuoFilter2.size()=" + listTuihuoFilter2.size());
//
//		// ���ι����쳣
//		if (hasNotXianxiang == false) {
//			int xianxiangId = GetXianxiangID(xianxiang);
//			System.out.println("���ι���!" + " �쳣=" + xianxiang + " xianxiangId=" + xianxiangId);
//			
//			ResultTuiHuo = 0;
//			for (Tuihuo item : listTuihuoFilter2) {
//
//				switch (xianxiangId) {
//
//				// δ�ҵ��쳣���� ,����
//				case 0:
//					jta1.append ("\n���ι����쳣 δ�ҵ��쳣���� ,����");
//					jta1.setBackground(Color.red);
//
//					System.out.println(" ���ι����쳣 δ�ҵ��쳣���� ,����");
//					ResultTuiHuo = 0;
//					return;
//
//				case 100:
//					ResultTuiHuo += item.getYichangCount();
//					break;
//		
//
//		
//
//		// listChuhuo = null;
//		// listTuihuo = null;
//
//	}


	protected void ReadDataToList(String filePath) throws Exception {
		// TODO Auto-generated method stub
		System.out.println("ReadDataToList!");
		 
		listInput =  ExcelReadKaoqing.exportListFromExcel(filePath);
//			listChuhuo = ExcelReadChuhuo.exportListFromExcel();
//			listTuihuo = ExcelReadTuihuo.exportListFromExcel();
		 

	}

}
