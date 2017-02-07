package linhao;

import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.filechooser.FileFilter;

public class View {

	static String filename,dirname;
	static JButton button1, button2, button3, button4, button5;
	static JTextField text1, text2;
	static JFileChooser jfc1,jfc2;
	static Model m;

	public static void createView() {//==================================================================界面呈现

		JLabel label1 = new JLabel("单个文件:");
		JLabel label2 = new JLabel("多个文件:");
		text1 = new JTextField();
		text2 = new JTextField();
		button1 = new JButton("打开");
		button2 = new JButton("选择");
		button3 = new JButton("开始转换");
		button4 = new JButton("清除");
		button5 = new JButton("帮助");
		m = new Model();

		button1.addActionListener(new ActionListener() {//======================================================按钮：单个文件
			@Override
			public void actionPerformed(ActionEvent e) {
				jfc1 = new JFileChooser();
				jfc1.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc1.setCurrentDirectory(new File("."));
				jfc1.showDialog(new JLabel(), "选择");
				text1.setText(jfc1.getSelectedFile().getAbsolutePath());
				filename = jfc1.getSelectedFile().getAbsolutePath();
				text2.setText("");
			}
		});

		button2.addActionListener(new ActionListener() {//======================================================按钮：多个文件
			@Override
			public void actionPerformed(ActionEvent e) {
				jfc2 = new JFileChooser();
				jfc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				jfc2.setCurrentDirectory(new File("."));
				jfc2.showDialog(new JLabel(), "选择");
				text2.setText(jfc2.getSelectedFile().getAbsolutePath());
				dirname = jfc2.getSelectedFile().getAbsolutePath();
				text1.setText("");
			}
		});

		button3.addActionListener(new ActionListener() {//======================================================按钮：开始转换
			@Override
			public void actionPerformed(ActionEvent e) {
				//====================================================================转换单个文件
				if(filename != null){
					try {
						m.readExcel(filename);
					} catch (Exception e1) {
						e1.printStackTrace();
					}
				}
				//====================================================================转换多个文件
				if (dirname != null) {
					File dir2 = new File(dirname);
					String[] allFiles=dir2.list();// 返回该目录下所有文件及文件夹数组
					for (int i = 0; i < allFiles.length; i++) {
						if(allFiles[i].endsWith(".xls")){
							try {
								m.readExcel(dir2+"\\"+allFiles[i]);
								//JOptionPane.showMessageDialog(null,dir+"\\"+allFiles[i]);
							} catch (Exception e1) {
								e1.printStackTrace();
							}
						}
					}
				}
				//====================================================================若无选择文件
				try {
					if(filename == null & dirname ==null){//没有选择任何文件
						JOptionPane.showMessageDialog(null,"您未选择任何文件");
					}else if(filename == null){//如果不是选择单个文件
						File dir2 = new File(dirname);
						Runtime.getRuntime().exec("cmd /c start explorer "+dir2);
						//JOptionPane.showMessageDialog(null,"转换完成");
						//JOptionPane.showMessageDialog(null,dir2);
					}else if(dirname ==null){//如果不是选择多个文件
						File dir1 = new File(filename);
						Runtime.getRuntime().exec("cmd /c start explorer "+dir1.getParent());
					}else{//既选择单个文件，又选择多个文件
						File dir2 = new File(dirname);
						Runtime.getRuntime().exec("cmd /c start explorer "+dir2);
						File dir1 = new File(filename);
						Runtime.getRuntime().exec("cmd /c start explorer "+dir1.getParent());
					}
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			}
		});
		
		button4.addActionListener(new ActionListener() {//======================================================按钮：清空选择
			@Override
			public void actionPerformed(ActionEvent e) {
				text1.setText("");
				text2.setText("");
			}
		});
		
		button5.addActionListener(new ActionListener() {//======================================================按钮：帮助信息
			@Override
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(null,"工具要求：\r\n1.客户信息表须在Excel首页\r\n2.表头为\"金额\"的仅有一列");
			}
		});

		JPanel panel = new JPanel();
		panel.setLayout(new GridLayout(3, 3));
		panel.add(label1);
		panel.add(text1);
		panel.add(button1);
		panel.add(label2);
		panel.add(text2);
		panel.add(button2);
		panel.add(button4);
		panel.add(button3);
		panel.add(button5);

		JFrame frame = new JFrame("惠来联社批量转格式工具v1.7");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(300, 120);
		frame.setContentPane(panel);
		frame.setVisible(true);

	}

	public static void main(String[] args) {
		// 单独线程管理UI
		createView();
	}

}
