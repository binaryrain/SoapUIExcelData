import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.xmlbeans.XmlException;

import com.eviware.soapui.impl.WsdlInterfaceFactory;
import com.eviware.soapui.impl.wsdl.WsdlInterface;
import com.eviware.soapui.impl.wsdl.WsdlOperation;
import com.eviware.soapui.impl.wsdl.WsdlProject;
import com.eviware.soapui.impl.wsdl.WsdlRequest;
import com.eviware.soapui.impl.wsdl.WsdlSubmit;
import com.eviware.soapui.impl.wsdl.WsdlSubmitContext;
import com.eviware.soapui.model.iface.Operation;
import com.eviware.soapui.model.iface.Request.SubmitException;
import com.eviware.soapui.model.iface.Response;
import com.eviware.soapui.support.SoapUIException;
import com.eviware.soapui.support.xml.XmlUtils;

public class SoapUIHelper {
	public static String wsdlUrl = "";
	public static String methodName = "";
	WsdlProject project;
	WsdlInterface iface;

	public SoapUIHelper(String wsdl) throws XmlException, IOException, SoapUIException {

		wsdlUrl = wsdl;
		if (wsdlUrl.toLowerCase().endsWith(".svc")) {
			wsdlUrl = wsdlUrl + "?singlewsdl";
		} else if (!wsdlUrl.toLowerCase().endsWith("?wsdl") && !wsdlUrl.toLowerCase().endsWith("?singlewsdl")) {
			wsdlUrl = wsdlUrl + "?wsdl";
		}
		project = new WsdlProject();
		iface = WsdlInterfaceFactory.importWsdl(project, wsdlUrl, true)[0];
	}

	public List<String> GetMethods() throws XmlException, IOException, SoapUIException {

		List<String> _methods = new ArrayList<String>();

		Operation[] ops = iface.getAllOperations();
		for (int i = 0; i < ops.length; i++) {
			_methods.add(ops[i].getName());
		}
		return _methods;
	}

	private void CreateSampleXml() throws IOException {
		// get desired operation
		WsdlOperation operation = (WsdlOperation) iface.getOperationByName(methodName);
		// generate the request content from the schema
		String xmlSample = operation.createRequest(true);
		java.io.FileWriter fw = new java.io.FileWriter(methodName + "__samplerequest" + ".xml");
		fw.write(xmlSample);
		fw.close();
	}

	public void CreateExcel() throws IOException, InterruptedException {
		CreateSampleXml();
		String appRootPath = System.getProperty("user.dir") + "\\";
		Process process = new ProcessBuilder(Arrays.asList(appRootPath + "XML2SoapXML.exe", SoapUIHelper.wsdlUrl,
				appRootPath + methodName + "__samplerequest" + ".xml")).start();
		int exitVal = process.waitFor();
		if (exitVal != 0) {
			throw new InterruptedException();
		}
	}

	public List<String> CreateRequests() throws IOException, InterruptedException {
		List<String> _reqs = new ArrayList<String>();
		String appRootPath = System.getProperty("user.dir") + "\\";
		Process process = new ProcessBuilder(Arrays.asList(appRootPath + "XML2SoapXML.exe", wsdlUrl,
				appRootPath + methodName + "__samplerequest" + ".xml", methodName + "_requests")).start();
		int exitVal = process.waitFor();
		if (exitVal != 0) {
			throw new InterruptedException();
		}

		File ff = new File(appRootPath + methodName + "_requests");

		File[] listOfFiles = ff.listFiles((d, n) -> n.endsWith(".xml"));
		for (int i = 0; i < listOfFiles.length; i++) {
			_reqs.add(listOfFiles[i].getName());
		}
		return _reqs;
	}

	public List<String> GetExistingRequests() throws IOException, InterruptedException {
		List<String> _reqs = new ArrayList<String>();
		String appRootPath = System.getProperty("user.dir") + "\\";

		File ff = new File(appRootPath + methodName + "_requests");

		File[] listOfFiles = ff.listFiles((d, n) -> n.endsWith(".xml"));
		for (int i = 0; i < listOfFiles.length; i++) {
			_reqs.add(listOfFiles[i].getName());
		}
		return _reqs;
	}

	public String CreateResponses(String requestXml) {
		String content = "";
		try {
			WsdlOperation operation = (WsdlOperation) iface.getOperationByName(methodName);
			File ff = new File(System.getProperty("user.dir") + "\\" + methodName + "_requests");

			File[] listOfFiles = ff.listFiles((d, n) -> n.equals(requestXml));

			Path path = Paths.get(System.getProperty("user.dir") + "\\" + methodName + "_responses");
			Files.createDirectories(path);

			for (int i = 0; i < listOfFiles.length; i++) {

				byte[] encoded = Files.readAllBytes(Paths.get(System.getProperty("user.dir") + "\\" + methodName
						+ "_requests" + "\\" + listOfFiles[i].getName()));
				String requestContent = new String(encoded, "UTF-8");

				WsdlRequest request = operation.addNewRequest("myRequest");
				request.setRequestContent(requestContent);

				WsdlSubmit<WsdlRequest> submit = request.submit(new WsdlSubmitContext(request), false);

				Response response = submit.getResponse();

				content = response.getContentAsString();
				if (content != null) {
					content = XmlUtils.prettyPrintXml(content);
					java.io.FileWriter fw = new java.io.FileWriter(System.getProperty("user.dir") + "\\" + methodName
							+ "_responses" + "\\" + listOfFiles[i].getName());
					fw.write(content);
					fw.close();
				} else {
					content = "Request data is incorrect. Error occured";
				}

			}
		} catch (IOException ex) {
			ex.printStackTrace();
		} catch (SubmitException ex) {
			ex.printStackTrace();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
		return content;
	}
}
