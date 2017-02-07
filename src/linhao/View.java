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

	public static void createView() {//==================================================================�������

		JLabel label1 = new JLabel("�����ļ�:");
		JLabel label2 = new JLabel("����ļ�:");
		text1 = new JTextField();
		text2 = new JTextField();
		button1 = new JButton("��");
		button2 = new JButton("ѡ��");
		button3 = new JButton("��ʼת��");
		button4 = new JButton("���");
		button5 = new JButton("����");
		m = new Model();

		button1.addActionListener(new ActionListener() {//======================================================��ť�������ļ�
			@Override
			public void actionPerformed(ActionEvent e) {
				jfc1 = new JFileChooser();
				jfc1.setFileSelectionMode(JFileChooser.FILES_ONLY);
				jfc1.setCurrentDirectory(new File("."));
				jfc1.showDialog(new JLabel(), "ѡ��");
				text1.setText(jfc1.getSelectedFile().getAbsolutePath());
				filename = jfc1.getSelectedFile().getAbsolutePath();
				text2.setText("");
			}
		});

		button2.addActionListener(new ActionListener() {//======================================================��ť������ļ�
			@Override
			public void actionPerformed(ActionEvent e) {
				jfc2 = new JFileChooser();
				jfc2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				jfc2.setCurrentDirectory(new File("."));
				jfc2.showDialog(new JLabel(), "ѡ��");
				text2.setText(jfc2.getSelectedFile().getAbsolutePath());
				dirname = jfc2.getSelectedFile().getAbsolutePath();
				text1.setText("");
			}
		});

		button3.addActionListener(new ActionListener() {//======================================================��ť����ʼת��
			@Override
			public void actionPerformed(ActionEvent e) {
				//====================================================================ת�������ļ�
				if(filename != null){
					try {
						m.readExcel(filename);
					} catch (Exception e1) {
						e1.printStackTrace();
					}
				}
				//====================================================================ת������ļ�
				if (dirname != null) {
					File dir2 = new File(dirname);
					String[] allFiles=dir2.list();// ���ظ�Ŀ¼�������ļ����ļ�������
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
				//====================================================================����ѡ���ļ�
				try {
					if(filename == null & dirname ==null){//û��ѡ���κ��ļ�
						JOptionPane.showMessageDialog(null,"��δѡ���κ��ļ�");
					}else if(filename == null){//�������ѡ�񵥸��ļ�
						File dir2 = new File(dirname);
						Runtime.getRuntime().exec("cmd /c start explorer "+dir2);
						//JOptionPane.showMessageDialog(null,"ת�����");
						//JOptionPane.showMessageDialog(null,dir2);
					}else if(dirname ==null){//�������ѡ�����ļ�
						File dir1 = new File(filename);
						Runtime.getRuntime().exec("cmd /c start explorer "+dir1.getParent());
					}else{//��ѡ�񵥸��ļ�����ѡ�����ļ�
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
		
		button4.addActionListener(new ActionListener() {//======================================================��ť�����ѡ��
			@Override
			public void actionPerformed(ActionEvent e) {
				text1.setText("");
				text2.setText("");
			}
		});
		
		button5.addActionListener(new ActionListener() {//======================================================��ť��������Ϣ
			@Override
			public void actionPerformed(ActionEvent e) {
				JOptionPane.showMessageDialog(null,"����Ҫ��\r\n1.�ͻ���Ϣ������Excel��ҳ\r\n2.��ͷΪ\"���\"�Ľ���һ��");
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

		JFrame frame = new JFrame("������������ת��ʽ����v1.7");
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setSize(300, 120);
		frame.setContentPane(panel);
		frame.setVisible(true);

	}

	public static void main(String[] args) {
		// �����̹߳���UI
		createView();
	}

}
