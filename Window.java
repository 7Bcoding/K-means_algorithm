
 
import java.awt.*;
import java.awt.event.*;
import javax.swing.table.DefaultTableModel;
import javax.swing.*;
import java.util.*;
import java.io.*;
import javax.swing.table.*;
import jlfx.UpdateExcel2003;
import java.util.*;
import jxl.*;
import java.lang.Math;
import java.io.File;  
 
public class Window extends JFrame{
  JFrame f_major = new JFrame("商务智能");
  JTabbedPane tp = new JTabbedPane();
 
  Font ft = new Font("Serif", Font.TRUETYPE_FONT, 18);
  Font ft1 = new Font("Serif", Font.ROMAN_BASELINE, 20);
  Font ft2 = new Font("Serif", Font.ROMAN_BASELINE, 15);
  Font ft3 = new Font("Serif", Font.TRUETYPE_FONT, 16);
 
  JPanel panel = new JPanel();
 
	//控件定义
  JLabel yssj_display=new JLabel("------原始数据显示区------", JLabel.CENTER);
  JLabel input_l=new JLabel("-------输入-------", JLabel.CENTER);
  JLabel jlgs_l=new JLabel("聚类个数", JLabel.CENTER);
  JLabel jgxsq_l=new JLabel("-------结果显示区-------", JLabel.CENTER);
  JButton yssjshow_b1=new JButton("显示原始数据");
  JButton jgshow_b1=new JButton("显示聚类后分析数据");
  JButton Cz_b=new JButton("重置");
  JTable jTable_1 = new   JTable();
  JTable jTable_2 = new   JTable();
  JComboBox num_cb1=new JComboBox();
  JScrollPane a1=new JScrollPane(jTable_1,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
	  JScrollPane a2=new JScrollPane(jTable_2,JScrollPane.VERTICAL_SCROLLBAR_ALWAYS,JScrollPane.HORIZONTAL_SCROLLBAR_ALWAYS);
 
class Flower{
     int n=0;
	 double x1=0;
     double x2=0;
	 double x3=0;
	 double x4=0; 
};
 Flower[] flower;
 int N;
 
public static void main(String[] args) {
 Window win = new Window();
 win.go();
}	  
  public void go() {
		  //窗体界面设计
	f_major.setSize(900,600);
	f_major.getContentPane().setLayout(new BorderLayout());
	f_major.getContentPane().add("Center", tp);
	
	f_major.setFont(ft);
    f_major.setVisible(true);
    f_major.setResizable(true);
 
    //选项卡界面设计
    tp.add("K-means聚类分析", panel);
    tp.setFont(ft);
 
    //界面设计
    panel.setLayout(null);
	//原始数据显示界面
    yssj_display.setSize(250, 40);
	yssj_display.setLocation(0, 0);
	yssj_display.setFont(ft1);
	      
    jTable_1.setSize(400, 350);
    jTable_1.setLocation(10, 40);
	      
     //输入界面
    input_l.setSize(185, 40);
	input_l.setLocation(0, 390);
	input_l.setFont(ft1);
	      
	jlgs_l.setSize(100, 40);
	jlgs_l.setLocation(10, 430);
	jlgs_l.setFont(ft2);
	      
	num_cb1.setSize(80, 40);
	num_cb1.setLocation(110, 430);
	num_cb1.addItem("3");
	      
	yssjshow_b1.setSize(120, 30);
	yssjshow_b1.setLocation(200,430);
   //重置
    Cz_b.setSize(150, 40);
	Cz_b.setLocation(550, 460);
	      
     //结果显示区界面
	  jgxsq_l.setSize(200, 40);
     jgxsq_l.setLocation(450, 0);
     jgxsq_l.setFont(ft1);
     
     jTable_2.setSize(400, 350);
     jTable_2.setLocation(450, 40);
	      
     a1.setViewportView(jTable_1);
     a1.setSize(400, 350);
     a1.setLocation(10, 40);
	      
     a2.setViewportView(jTable_2);
     a2.setSize(400, 350);
     a2.setLocation(450, 40);
    
     jgshow_b1.setSize(180, 30);
     jgshow_b1.setLocation(680, 390);
	      
     panel.add(yssj_display);
     panel.add(a1);
     panel.add(a2);
//   panel.add(jTable_1);
     panel.add(input_l);
     panel.add(jlgs_l);
     panel.add(num_cb1);
     panel.add(jgxsq_l);
//   panel.add(jTable_2);
     panel.add(jgshow_b1);
     panel.add(  yssjshow_b1);
     panel.add(Cz_b);
	
     //定义原始数据监控器
     yssjshow_b1.addActionListener(new ShowData());
	 //定义结果显示监控器
    jgshow_b1.addActionListener(new ShowSequence());      
	  }	  
     //ShowData内部类
class ShowData  implements ActionListener{
  public void actionPerformed(ActionEvent arg0) {
	 	try {  				 
	  InputStream stream = new FileInputStream(new File("jlfx_data.xls"));
	  Workbook rwb = Workbook.getWorkbook(stream);
		Cell cell = null;
	    Sheet sheet = rwb.getSheet(0);
		DefaultTableModel tableModel=(DefaultTableModel) jTable_1.getModel();
		tableModel.setColumnCount(5);
		tableModel.setRowCount(sheet.getRows());
		Object[] object=new Object[jTable_1.getColumnCount()];
			 N=sheet.getRows()-1;
	if(sheet.getRows()>2){
		flower=new Flower[sheet.getRows()];
		 for(int i = 0; i < sheet.getRows(); i++){
			flower[i] = new Flower();
		}
		  	for(int i=0;i<sheet.getRows();i++){
		     for(int j=0;j<5;j++)
	          {
		       cell=sheet.getCell(j,i);
		  	   if(cell.getType()==CellType.LABEL)
		       {	
	  		   LabelCell labelcell=(LabelCell)cell;
	  		   object[j]=labelcell.getString();
               jTable_1.setValueAt(labelcell.getString(),i,j);
			 }
	else if(cell.getType()==CellType.NUMBER)
		 {
	        Double numd;
		    double numi;
		     NumberCell numc10=(NumberCell)cell;
	         numd=new Double(numc10.getValue());
			 numi=numd.doubleValue();
			 object[j]=numi;
			 jTable_1.setValueAt(numi,i,j);
			 if(j==0)flower[i].n=(int)numi;
			 else if(j==1)flower[i].x1=numi;
			 else if(j==2)flower[i].x2=numi;
			 else if(j==3)flower[i].x3=numi;
	         else if(j==4)flower[i].x4=numi;
			}
			  	}					  	
}					  	rwb.close(); 
 
                                }
}
catch  (Exception e)  
	    {  
	     System.out.println(e);  
         }   
		}
					}
			//ShowSwquence内部类
	class ShowSequence implements ActionListener{
		public void actionPerformed(ActionEvent arg0) {
			// 获取聚类个数			
			int k=Integer.parseInt(String.valueOf( num_cb1.getSelectedItem()));
			double[][] box=new double[k+1][N+1];
			int[][] count ={{0},{0},{0},{0}};
			int[][] count_temp ={{0},{0},{0},{0}}; 
			double [][] center = new double[k+1][5];
			//初始中心放在一个三行四列的数组里
			center[1][1] = flower[1].x1;
			center[1][2] = flower[1].x2;
			center[1][3] = flower[1].x3;
			center[1][4] = flower[1].x4;
			center[2][1] = flower[51].x1;
			center[2][2] = flower[51].x2;
			center[2][3] = flower[51].x3;
			center[2][4] = flower[51].x4;
			center[3][1] = flower[101].x1;
			center[3][2] = flower[101].x2;
			center[3][3] = flower[101].x3;
			center[3][4] = flower[101].x4;		
			double[] E={1.0,0,0,0};
			double[][] sum_temp={{0,0,0,0,0},{0,0,0,0,0},{0,0,0,0,0},{0,0,0,0,0}};
			double x=0.0, temp;
			double[] d={0,0,0,0};
			int tag=1;
			int num=0;
			//double[] sum = {0,0,0,0};
					
			while(Math.abs(E[0] - x) > 0.0001)
			{
				num++;
				x = E[0];
				for(int i=1; i<4; i++)count[i][0] = 0;
				for(int f=1; f<151; f++)
				{
					//求出所有数据与聚类中心均值的差值存放在d[j]中
				for(int j=1; j<4; j++)
				{
					
				d[j] = Math.pow(Math.abs(center[j][1] - flower[f].x1), 2);
				d[j] += Math.pow(Math.abs(center[j][2] - flower[f].x2), 2);
				d[j] += Math.pow(Math.abs(center[j][3] - flower[f].x3), 2);
				d[j] += Math.pow(Math.abs(center[j][4] - flower[f].x4), 2);
				d[j] = Math.sqrt(d[j]);
					
				}
				
				temp = d[1];
				tag=1;
				//求出每一行数据与聚类中心均值的欧式距离的最小值，并归于该聚类中心
				for(int m=1; m<4; m++)
				{		
						if(d[m] < temp)
						{
							temp = d[m];
							tag = m;//tag标记数据分类到相应的聚类中心
													}
						}
			//每找到一行数据的归属，就将相应的聚类中心的数据行数+1
			count[tag][0] = count[tag][0] + 1;
		    //将某一行数据归于相应的聚类中心的那一行
			box[tag][count[tag][0]] = f; 
			E[tag] += Math.pow(d[tag],2);
	               //算出每一列的数据的总和
	     	sum_temp[tag][1] += flower[f].x1;
     		sum_temp[tag][2] += flower[f].x2；
			sum_temp[tag][3] += flower[f].x3;
			sum_temp[tag][4] += flower[f].x4;			
			}
			E[0] = E[1]+E[2]+E[3];
 
			for(int i=1; i<4; i++)
		    //3个聚类中心
			{
				for(int j=1; j<=4; j++)
			//4列数据
			{
//将每一列数据总和除以每一列的数据的行数，
  //得到相应那一列的数据的均值，4列数据的均值便是新的聚类中心
			sum_temp[i][j] = sum_temp[i][j]/count[i][0];
		//更新了新的聚类中心，赋值给center数组，供下一轮循环使用
			center[i][j] = sum_temp[i][j];
			sum_temp[i][j] = 0;
				}
							
			}					
			for(int i=1; i<4; i++)
			{
				E[i] = 0;
				d[i] = 0;
			count_temp[i][0] = count[i][0];
			}
						
					}				
		 System.out.println("循环次数："+num);		
			//分类完成,输出每个聚类条数；
        for(int i=1; i<4; i++)					
          {
			System.out.println("第"+i+"个聚类项数："+count_temp[i][0]);
		}
		//输出全部聚类内容；
		               for(int i=1; i<4; i++)
			    {
				for(int j=1; j<=count[i][0]; j++)
				{
				try{
		File file = new File("jieguo.xls");
				InputStream stream = new FileInputStream(new File("jieguo.xls"));			// 获取Excel文件对象
               Workbook rwb = Workbook.getWorkbook(stream);
	UpdateExcel.updateExcel(file, 0, (i-1)*4+0, j,flower[(int)box[i][j]].x1);
				UpdateExcel.updateExcel(file, 0, (i-1)*4+1, j,flower[(int)box[i][j]].x2);
				UpdateExcel.updateExcel(file, 0, (i-1)*4+2, j,flower[(int)box[i][j]].x3);
				UpdateExcel.updateExcel(file, 0, (i-1)*4+3, j,flower[(int)box[i][j]].x4);
			rwb.close();
		}catch  (Exception e)  
		  {  				
	           System.out.println(e);  
				 } 
				}
				}
			try {  
					 
            InputStream stream = new FileInputStream(new File("jieguo.xls"));
		    // 获取excel文件对象
		    Workbook rwb = Workbook.getWorkbook(stream);
		    Cell cell = null;
		    Sheet sheet = rwb.getSheet(0);
			DefaultTableModel tableModel=(DefaultTableModel) jTable_2.getModel();
			tableModel.setColumnCount(sheet.getColumns());
			tableModel.setRowCount(sheet.getRows());
		    Object[] object=new Object[jTable_2.getColumnCount()];
			 for(int q=0;q<sheet.getRows();q++)
				{//列循环
				 for(int p=0;p<sheet.getColumns();p++)
				  {
					 cell=sheet.getCell(p,q);
					 if(cell.getType()==CellType.LABEL)
					  	 {	
					  	  LabelCell labelcell=(LabelCell)cell;
					  	  object[p]=labelcell.getString();
					  	  jTable_2.setValueAt(labelcell.getString(),q,p);   		  							         }
		             else if(cell.getType()==CellType.NUMBER)
					      {
					  	  Double numd;
					  	  double numi;
						  NumberCell numc10=(NumberCell)cell;
						  numd=new Double(numc10.getValue());
						  numi=numd.doubleValue();
						  object[p]=numi;
						  jTable_2.setValueAt(numi,q,p);	
					  	   }
					   }
			 }		  			
					  			rwb.close(); 
				}
	           catch  (Exception e)  
	           {  
	            System.out.println(e);  
	            }   
				}
				}
		}
