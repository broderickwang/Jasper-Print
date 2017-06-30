package Jasper;

import JasperUtil.PrintType;
import net.sf.jasperreports.engine.*;
import net.sf.jasperreports.engine.base.JRBaseReport;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jasperreports.engine.export.JRPdfExporter;
import net.sf.jasperreports.engine.export.JRRtfExporter;
import net.sf.jasperreports.engine.export.JRXlsExporter;
import net.sf.jasperreports.engine.type.OrientationEnum;
import net.sf.jasperreports.engine.util.FileBufferedOutputStream;
import net.sf.jasperreports.engine.util.JRLoader;
import net.sf.jasperreports.export.HtmlExporterConfiguration;
import org.apache.commons.collections.map.HashedMap;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.sql.Connection;
import java.util.Collection;
import java.util.Map;

/*
 *Created by Broderick
 * User: Broderick
 * Date: 2017/6/29
 * Time: 10:41
 * Version: 1.0
 * Description:
 * Email:wangchengda1990@gmail.com
**/
public class JRHelper {
    private PrintType type;

    private String outName;

    private File jasperFile;

    private Map paramers;

    private Connection connection;

    private Collection datas;

    private HttpServletResponse response;

    private JRHelper(Builder builder) {
        this.type = builder.type;
        this.outName = builder.outName;
        this.jasperFile = builder.jasperFile;
        this.paramers = builder.paramers;
        this.connection = builder.connection;
        this.datas = builder.datas;
        this.response = builder.response;
    }

    private void export(){

        try {
            JasperReport jasperReport = (JasperReport) JRLoader.loadObject(jasperFile);
            prepareReport(jasperReport, type);

            JasperPrint jasperPrint;
            if(connection == null) {
                JRDataSource ds = new JRBeanCollectionDataSource(datas, false);
                jasperPrint = JasperFillManager.fillReport(jasperReport, paramers, ds);
            } else
                jasperPrint = JasperFillManager.fillReport(jasperReport,paramers,connection);

            switch (type){
                case PDF_TYPE:
                    exportPdfFile(jasperPrint,outName, response);
                    break;
                case PDF_IO_TYPE:
                    exprotPdfIO(jasperPrint,response);
                    break;
                case HTML_TYPE:
                    exportHtml(jasperPrint,outName, response);
                    break;
                case EXCEL_TYPE:
                    exportExcel(jasperPrint,outName, response);
                    break;
                case WORD_TYPE:
                    exportWord(jasperPrint,outName, response);
                    break;
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * 如果导出的是excel，则需要去掉周围的margin
     * @param jasperReport
     * @param type
     */
    private void prepareReport(JasperReport jasperReport, PrintType type) {
        switch (type){
            case EXCEL_TYPE:
                try {
                    Field margin = JRBaseReport.class
                            .getDeclaredField("leftMargin");
                    margin.setAccessible(true);
                    margin.setInt(jasperReport, 0);
                    margin = JRBaseReport.class.getDeclaredField("topMargin");
                    margin.setAccessible(true);
                    margin.setInt(jasperReport, 0);
                    margin = JRBaseReport.class.getDeclaredField("bottomMargin");
                    margin.setAccessible(true);
                    margin.setInt(jasperReport, 0);
                    Field pageHeight = JRBaseReport.class
                            .getDeclaredField("pageHeight");
                    pageHeight.setAccessible(true);
                    pageHeight.setInt(jasperReport, 2147483647);
                } catch (Exception exception) {
                }
                break;
        }

    }

    /**
     * 导出excel
     * @param jasperPrint
     * @param defaultFilename
     * @param response
     * @throws IOException
     * @throws JRException
     */
    private void exportExcel(JasperPrint jasperPrint, String defaultFilename,
                             HttpServletResponse response) throws IOException, JRException {

        String defaultname=null;
        if(defaultFilename.trim()!=null&&defaultFilename!=null){
            defaultname=defaultFilename+".xls";
        }else{
            defaultname="export.xls";
        }
        String fileName = new String(defaultname.getBytes("gbk"), "utf-8");
        System.out.println("1"); // only this line is printed in console not others
        JRXlsExporter exporter1=null;
        try{
            exporter1 = new  JRXlsExporter();
            System.out.println("2");
        }
        catch(Exception e){
            e.printStackTrace();
        }
        System.out.println("3");
        response.setContentType("application/xls");
        System.out.println("4");
        response.setHeader("Content-Disposition", "inline; filename=\""+fileName+"\"");
        System.out.println("5");
        exporter1.setParameter(JRExporterParameter.JASPER_PRINT,  jasperPrint);
        System.out.println("6");
        exporter1.setParameter(JRExporterParameter.OUTPUT_STREAM,  response.getOutputStream());
        System.out.println("7");
        exporter1.exportReport();
    }

    /**
     * 导出pdf文件
     * 注意此处中文问题，

     * 这里应该详细说：主要在studio里变下就行了。

     * 下边的设置就在点字段的属性后出现。
     * pdf font name ：STSong-Light ，pdf encoding ：UniGB-UCS2-H
     */
    private void exportPdfFile(JasperPrint jasperPrint, String defaultFilename,
                               HttpServletResponse response) throws IOException, JRException {
        response.setContentType("application/pdf");
        String defaultname=null;
        if(defaultFilename.trim()!=null&&defaultFilename!=null){
            defaultname=defaultFilename+".pdf";
        }else{
            defaultname="export.pdf";
        }
        String fileName = new String(defaultname.getBytes("GBK"), "ISO8859_1");
        response.setHeader("Content-disposition", "attachment; filename="
                + fileName);
        ServletOutputStream ouputStream = response.getOutputStream();
        JasperExportManager.exportReportToPdfStream(jasperPrint,
                ouputStream);
        ouputStream.flush();
        ouputStream.close();
    }

    /**
     * 导出PDF，直接显示在servlet web页面
     * @param jasperPrint
     * @param response
     * @throws IOException
     */
    private void exprotPdfIO(JasperPrint jasperPrint,HttpServletResponse response)throws IOException{
        FileBufferedOutputStream fbos = new FileBufferedOutputStream();
        JRPdfExporter exporter = new JRPdfExporter();
        exporter.setParameter(JRExporterParameter.OUTPUT_STREAM, fbos);
        exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
        try {
            exporter.exportReport();
            fbos.close();
            if (fbos.size() > 0) {
                //response.setContentType("application/pdf;charset=utf-8");
                response.setContentType("application/pdf");
                response.setContentLength(fbos.size());
                ServletOutputStream ouputStream = response.getOutputStream();
                try {
                    fbos.writeData(ouputStream);
                    fbos.dispose();
                    ouputStream.flush();
                } finally {
                    if (null != ouputStream) {
                        ouputStream.close();
                    }
                }
            }
        } catch (JRException e1) {
            e1.printStackTrace();
        }finally{
            if(null !=fbos){
                fbos.close();
                fbos.dispose();
            }
        }
    }

    /**
     * 导出html
     * @param jasperPrint
     * @param defaultFilename
     * @param response
     * @throws IOException
     * @throws JRException
     */
    private void exportHtml(JasperPrint jasperPrint,String defaultFilename,
                            HttpServletResponse response) throws IOException, JRException {
        response.setContentType("text/html");
        //强制设置分页，不再使用过时的方法
        DefaultJasperReportsContext context = DefaultJasperReportsContext.getInstance();
        context.setProperty(HtmlExporterConfiguration.PROPERTY_BETWEEN_PAGES_HTML,"<DIV STYLE='page-break-before:always;'></DIV>");

        jasperPrint.setOrientation(OrientationEnum.LANDSCAPE);

        String defaultname=null;
        if(defaultFilename.trim()!=null&&defaultFilename!=null){
            defaultname=defaultFilename+".html";
        }else{
            defaultname="export.html";
        }

        JasperExportManager.getInstance(context).exportReportToHtmlFile(jasperPrint, defaultname);

        String[] names = defaultname.split("\\/");

        String name = names[names.length-1];

        response.sendRedirect(name);
    }

    /**
     * 导出word
     */
    private void exportWord(JasperPrint jasperPrint,String defaultFilename,
                            HttpServletResponse response)
            throws JRException, IOException {
        response.setContentType("application/msword;charset=utf-8");
        String defaultname=null;
        if(defaultFilename.trim()!=null&&defaultFilename!=null){
            defaultname=defaultFilename+".doc";
        }else{
            defaultname="export.doc";
        }
        String fileName = new String(defaultname.getBytes("GBK"), "utf-8");
        response.setHeader("Content-disposition", "attachment; filename="
                + fileName);
        JRExporter exporter = new JRRtfExporter();
        exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
        exporter.setParameter(JRExporterParameter.OUTPUT_STREAM, response
                .getOutputStream());

        exporter.exportReport();
    }

    /**
     * JRHelper 的构建类，JRHelper采用链式调用的方式进行Jasper文件的打印
     */
    public static class Builder{
        private PrintType type;

        private String outName;

        private File jasperFile;

        private Map paramers;

        private Connection connection;

        private Collection datas;

        private HttpServletResponse response;

        public Builder() {
            paramers = new HashedMap();
        }

        /**
         * 设置打印类型
         * @param type  打印类型
         * @return      Builder
         */
        public Builder type(PrintType type){
            this.type = type;
            return this;
        }

        /**
         * 设置输出文件名称，HTML要特别主要，因为输出之后需要定向到HTML，所以要带路径
         * @param outName   文件名称
         * @return          Builder
         */
        public Builder outName(String outName){
            this.outName = outName;
            return this;
        }

        /**
         * 设置Jasper模板文件
         * @param jasperFile    模板文件
         * @return              Builder
         */
        public Builder jasperFile(File jasperFile){
            this.jasperFile = jasperFile;
            return this;
        }

        /**
         * 添加Jasper模板需要的参数，参数内部以map的形式存储
         * @param key       参数的key
         * @param value     参数的值
         * @return          Builder
         */
        public Builder addP(String key, Object value){
            paramers.put(key,value);
            return this;
        }

        /**
         * 设置Jasper模板内部的数据库连接，如果模板不是使用的数据库连接的方式，
         * 而是使用的Collection的方式，该方法可以不设置
         * @param connection    数据库连接
         * @return              Builder
         */
        public Builder connection(Connection connection){
            this.connection = connection;
            return this;
        }

        /**
         * 设置Jasper模板的数据集，如果模板采用的是数据库连接的方式，该方法可以不设置
         * @param collection    数据集
         * @return              Builder
         */
        public Builder collection(Collection collection){
            this.datas = collection;
            return this;
        }

        /**
         * 设置response
         * @param response  response
         * @return          Builder
         */
        public Builder response(HttpServletResponse response){
            this.response = response;
            return this;
        }

        public JRHelper build(){
            if(connection==null && (datas==null || datas.size()==0)){
                throw new IllegalArgumentException("请设置一种数据填充方式!调用collection()或connection()方法");
            }else if(connection!=null && (datas!=null && datas.size()>0))
                throw new IllegalArgumentException("至多只能设置一种数据填充方式!");
            else if(type == null)
                throw new IllegalArgumentException("请设置输出类型");
            else if(jasperFile == null)
                throw new IllegalArgumentException("请设置模板文件");
            else if(response == null)
                throw new IllegalArgumentException("请设置response");
            JRHelper helper = new JRHelper(this);
            helper.export();
            return helper;
        }
    }
}
