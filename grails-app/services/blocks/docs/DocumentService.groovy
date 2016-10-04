package blocks.docs

import fr.opensagres.xdocreport.core.document.SyntaxKind
import fr.opensagres.xdocreport.document.IXDocReport
import fr.opensagres.xdocreport.document.registry.XDocReportRegistry
import fr.opensagres.xdocreport.template.IContext
import fr.opensagres.xdocreport.template.TemplateEngineKind
import fr.opensagres.xdocreport.template.formatter.FieldsMetadata
import grails.transaction.Transactional
//import org.apache.commons.io.IOUtils
import org.apache.commons.logging.LogFactory
import org.artofsolving.jodconverter.OfficeDocumentConverter
import org.artofsolving.jodconverter.document.DefaultDocumentFormatRegistry
import org.artofsolving.jodconverter.document.DocumentFormatRegistry
import org.artofsolving.jodconverter.office.DefaultOfficeManagerConfiguration
import org.artofsolving.jodconverter.office.OfficeConnectionProtocol
import org.artofsolving.jodconverter.office.OfficeManager

import javax.annotation.PostConstruct
import javax.annotation.PreDestroy
import java.util.Map.Entry

@Transactional
class DocumentService {

    private static final log = LogFactory.getLog(this)

    def static grailsApplication

    private DocumentFormatRegistry formatRegistry
    private OfficeManager officeManager
    private OfficeDocumentConverter converter
    private boolean started

    @PostConstruct
    def init() {
        try{
            //lazy restart
            log.info("initializing document service");
            preDestroy();
            log.info("office path is $grailsApplication.config.officePath");
            officeManager = new DefaultOfficeManagerConfiguration()
                    .setOfficeHome(grailsApplication.config.officePath)
                    .setConnectionProtocol(OfficeConnectionProtocol.SOCKET)
                    .setPortNumber(8100)
                    .setTaskExecutionTimeout(30000)
                    .buildOfficeManager();
            formatRegistry = new DefaultDocumentFormatRegistry();
            converter = new OfficeDocumentConverter(officeManager, formatRegistry);
            officeManager.start();
            started = true;

        }catch(final Exception ex){
            log.error("Error while constructing converter. Skipped initialization.", ex);
        }
    }

    @PreDestroy
    public void preDestroy() {
        try{

            if ( officeManager != null) {
                officeManager.stop();
            }
        }catch(final Exception ex){
            log.error("Error while destroying converter. Skipped destruction.", ex);
        }
    }

    public void generateODT(InputStream inputStream, Map<String, Object> context, OutputStream outputStream)  {
        generateODT(inputStream, null, context, outputStream);
    }

    public void generateODT(InputStream inputStream, InputStream stylesStream, Map<String, Object> context, OutputStream outputStream) {
        try {

            log.info("generateODT("+context+")");
            IXDocReport report = XDocReportRegistry.getRegistry().loadReport(inputStream, TemplateEngineKind.Velocity);

            FieldsMetadata metadata = report.createFieldsMetadata();
            IContext xcontext = report.createContext();
            xcontext.put("htmlTemplate","");
            xcontext.put("wikiTemplate","");
            if (stylesStream != null) {
                xcontext.put("additionalStyles", new String(IOUtils.toByteArray(stylesStream), "UTF-8"));
            }
            for(Entry<String, Object> entry : context.entrySet()){
                xcontext.put(entry.getKey(), entry.getValue());
            }

            //xcontext.put("include", includeHelper);
            metadata.addFieldAsTextStyling("wikiTemplate", SyntaxKind.GWiki, true);
            metadata.addFieldAsTextStyling("htmlTemplate", SyntaxKind.Html, true);
            metadata.addFieldAsTextStyling("odtTemplate", SyntaxKind.NoEscape, true);
            metadata.addFieldAsTextStyling("rawTemplate", SyntaxKind.NoEscape, false);


            //for(String method : ClassWrapperRegistry.getHtmlMethods()){
            //	metadata.addFieldAsTextStyling(method, SyntaxKind.Html, true);
            //}

            metadata.addFieldAsList("tableItem");

            report.process(xcontext, outputStream);
        } catch (Exception e) {
            log.error("Processing document failed", e);
        }
    }

    def convertToPDF(InputStream odtInput, OutputStream pdfOutput) {
        if(!started){
            init();
        }
        try {
            File input = File.createTempFile("jodconveter-in", ""+System.currentTimeMillis());
            FileOutputStream outputStream = new FileOutputStream(input);
            try{
                IOUtils.copy(odtInput, outputStream);
            }finally{
                outputStream.close();
            }

            File output = File.createTempFile("jodconveter-out", ""+System.currentTimeMillis());
            converter.convert(input, output, formatRegistry.getFormatByExtension("pdf"));
            FileInputStream inputStream = new FileInputStream(output);
            try{
                IOUtils.copy(inputStream, pdfOutput);
            }finally{
                inputStream.close();
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
