import java.io.File;
import java.io.IOException;

import org.eclipse.swt.SWT;
import org.eclipse.swt.custom.StyleRange;
import org.eclipse.swt.custom.StyledText;
import org.eclipse.swt.layout.GridData;
import org.eclipse.swt.layout.GridLayout;
import org.eclipse.swt.widgets.Button;
import org.eclipse.swt.widgets.Display;
import org.eclipse.swt.widgets.Event;
import org.eclipse.swt.widgets.List;
import org.eclipse.swt.widgets.Listener;
import org.eclipse.swt.widgets.MessageBox;
import org.eclipse.swt.widgets.Shell;
import org.eclipse.swt.widgets.Text;

public class Program {

	Display display = new Display();
	Shell shell = new Shell(display, SWT.SHELL_TRIM | SWT.CENTER);
	SoapUIHelper objSoapUIHelper;
	Button btnChooseMethod;
	Button btnCreateExcel;
	Button btnCreateXml;
	Button btnHitService;
	Text txtWsdlUrl;
	List listMethods;
	List listXMLFiles;
	StyledText txtResponse;

	public Program() {
		try {
			InitUI();
			AssignShowServiceMethodsButtonHandler();
			AssignCreateExcelButtonHandler();
			AssignCreateRequestButtonHandler();
			AssignCallServiceButtonHandler();
			shell.setSize(850, 800);
			// shell.pack();
			shell.open();

			// Set up the event loop.

			while (!shell.isDisposed()) {
				if (!display.readAndDispatch()) {
					// If no more entries in event queue
					display.sleep();
				}
			}

		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			CleanUIResources();
		}
	}

	public static void main(String[] args) {
		new Program();
	}

	private void InitUI() {
		GridLayout layout = new GridLayout(4, true);
		layout.verticalSpacing = 2;
		// setLayout(layout);
		shell.setLayout(layout);

		// row 1 column first 3
		txtWsdlUrl = new Text(shell, SWT.None);
		GridData gdTxtWsdlUrl = new GridData(SWT.FILL, SWT.TOP, false, false);
		gdTxtWsdlUrl.heightHint = 30;
		gdTxtWsdlUrl.horizontalSpan = 3;
		txtWsdlUrl.setBounds(0, 0, 700, 30);
		txtWsdlUrl.setLayoutData(gdTxtWsdlUrl);

		// row 1 column 4
		btnChooseMethod = new Button(shell, SWT.None);
		GridData gdbtnChooseMethod = new GridData(SWT.FILL, SWT.TOP, false, false);
		gdbtnChooseMethod.heightHint = 30;
		btnChooseMethod.setText("Go");
		btnChooseMethod.setLayoutData(gdbtnChooseMethod);

		// row 2,3 column 1
		listMethods = new List(shell, SWT.BORDER | SWT.SINGLE | SWT.V_SCROLL);
		GridData gdlistMethods = new GridData(SWT.FILL, SWT.FILL, true, true);
		gdlistMethods.horizontalSpan = 1;
		gdlistMethods.verticalSpan = 2;
		listMethods.setLayoutData(gdlistMethods);

		// row 2 column 2
		btnCreateExcel = new Button(shell, SWT.None);
		GridData gdbtnCreateExcel = new GridData(SWT.FILL, SWT.TOP, false, false);
		gdbtnCreateExcel.heightHint = 30;
		btnCreateExcel.setText("Create Excel");
		btnCreateExcel.setLayoutData(gdbtnCreateExcel);

		// row 2 column 3
		btnCreateXml = new Button(shell, SWT.None);
		GridData gdbtnCreateXml = new GridData(SWT.FILL, SWT.TOP, false, false);
		gdbtnCreateXml.heightHint = 30;
		btnCreateXml.setText("Create Requests");
		btnCreateXml.setLayoutData(gdbtnCreateXml);

		// row 2 column 4
		btnHitService = new Button(shell, SWT.None);
		GridData gdbtnbtnHitService = new GridData(SWT.FILL, SWT.TOP, false, false);
		gdbtnbtnHitService.heightHint = 30;
		btnHitService.setText("Call Service");
		btnHitService.setLayoutData(gdbtnbtnHitService);

		// row 3 column 2
		listXMLFiles = new List(shell, SWT.FILL | SWT.SINGLE | SWT.V_SCROLL);
		GridData gdlistXMLFiles = new GridData(SWT.FILL, SWT.FILL, true, true);
		listXMLFiles.setLayoutData(gdlistXMLFiles);

		// row 3 column 3,4
		txtResponse = new StyledText(shell, SWT.MULTI | SWT.BORDER | SWT.WRAP | SWT.V_SCROLL);
		GridData gdtxtResponse = new GridData(GridData.FILL_BOTH); // new
																	// GridData(SWT.FILL,
																	// SWT.FILL,
																	// false,
																	// false);
		gdtxtResponse.horizontalSpan = 2;
		txtResponse.setLayoutData(gdtxtResponse);

	}

	private void AssignShowServiceMethodsButtonHandler() {
		// show methods
		btnChooseMethod.addListener(SWT.Selection, new Listener() {
			@Override
			public void handleEvent(Event arg0) {
				btnChooseMethod.setText("Processing...");
				btnChooseMethod.setEnabled(false);
				display.asyncExec(new Runnable() {
					@Override
					public void run() {
						try {
							objSoapUIHelper = new SoapUIHelper(txtWsdlUrl.getText().trim());
							listMethods.removeAll();
							java.util.List<String> allMethods = objSoapUIHelper.GetMethods();
							for (int i = 0, n = allMethods.size(); i < n; i++) {
								listMethods.add(allMethods.get(i));
							}
						} catch (Exception e) {
							MessageBox messageBox = new MessageBox(shell, SWT.ERROR);
							messageBox.setText("Error: Getting service methods");
							messageBox.setMessage("Verify service WSDL URL is correct.");
							messageBox.open();
						} finally {
							btnChooseMethod.setText("Go");
							btnChooseMethod.setEnabled(true);
						}
					}
				});
			}
		});
	}

	private void AssignCreateExcelButtonHandler() {
		// Create excel
		btnCreateExcel.addListener(SWT.Selection, new Listener() {

			@Override
			public void handleEvent(Event arg0) {
				btnCreateExcel.setText("Processing...");
				btnCreateExcel.setEnabled(false);
				display.asyncExec(new Runnable() {
					@Override
					public void run() {
						try {
							int selIndex = listMethods.getSelectionIndex();
							if (selIndex < 0) {
								MessageBox messageBox = new MessageBox(shell, SWT.ERROR);
								messageBox.setText("Select a method");
								messageBox.setMessage("Select method for which data needs to be created.");
								messageBox.open();
								return;
							}
							String dfdsd = listMethods.getItem(selIndex);
							SoapUIHelper.methodName = dfdsd;

							String appRootPath = System.getProperty("user.dir");
							File ff = new File(appRootPath);
							File[] listOfFiles = ff.listFiles((d, n) -> n.endsWith(".xlsx") || n.endsWith(".cs")
									|| n.endsWith(".xsd") || n.endsWith(".xml"));
							for (int i = 0; i < listOfFiles.length; i++) {
								File curFile = listOfFiles[i];
								curFile.delete();
							}
							objSoapUIHelper.CreateExcel();
							MessageBox messageBox = new MessageBox(shell, SWT.OK);
							messageBox.setText("Excel creation successful");
							messageBox.setMessage("Excel created at path : " + appRootPath);
							messageBox.open();
						} catch (IOException e) {
							MessageBox messageBox = new MessageBox(shell, SWT.ERROR);
							messageBox.setText("Sample Xml creation error.");
							messageBox.setMessage("Program failed to execute.");
							messageBox.open();
						} catch (InterruptedException e) {
							MessageBox messageBox = new MessageBox(shell, SWT.ERROR);
							messageBox.setText("Excel creation error.");
							messageBox.setMessage("Error while creating excel.");
							messageBox.open();
						} finally {
							btnCreateExcel.setText("Create Excel");
							btnCreateExcel.setEnabled(true);
						}
					}
				});
			}
		});
	}

	private void AssignCreateRequestButtonHandler() {
		// create requests
		btnCreateXml.addListener(SWT.Selection, new Listener() {
			@Override
			public void handleEvent(Event arg0) {
				btnCreateXml.setText("Processing...");
				display.asyncExec(new Runnable() {
					@Override
					public void run() {
						try {
							String dfdsd = listMethods.getItem(listMethods.getSelectionIndex());
							SoapUIHelper.methodName = dfdsd;
							java.util.List<String> allReqs = objSoapUIHelper.CreateRequests();
							listXMLFiles.removeAll();
							for (int i = 0, n = allReqs.size(); i < n; i++) {
								listXMLFiles.add(allReqs.get(i));
							}
						} catch (IOException e) {
							MessageBox messageBox = new MessageBox(shell, SWT.ERROR);
							messageBox.setText("Request Xml creation error.");
							messageBox.setMessage("Program failed to execute.");
							messageBox.open();
						} catch (InterruptedException e) {
							MessageBox messageBox = new MessageBox(shell, SWT.ERROR);
							messageBox.setText("Requests creation error.");
							messageBox.setMessage("Error while creating requests.");
							messageBox.open();
						} finally {
							btnCreateXml.setText("Create Requests");
						}
					}
				});
			}

		});
	}

	private void AssignCallServiceButtonHandler() {
		// call service
		btnHitService.addListener(SWT.Selection, new Listener() {
			@Override
			public void handleEvent(Event arg0) {
				txtResponse.setText("Processing...");
				display.asyncExec(new Runnable() {

					@Override
					public void run() {

						try {
							String dfdsd = listMethods.getItem(listMethods.getSelectionIndex());
							String requestFileName = listXMLFiles.getItem(listXMLFiles.getSelectionIndex());
							SoapUIHelper.methodName = dfdsd;
							String content = objSoapUIHelper.CreateResponses(requestFileName);
							txtResponse.setText(content);
							// editorPane.setText(content);
							java.util.List<XmlRegion> regions = new XmlRegionAnalyzer().analyzeXml(content);
							for (XmlRegion xr : regions) {
								int regionLength = xr.getEnd() - xr.getStart();
								StyleRange styleRange = new StyleRange();
								switch (xr.getXmlRegionType()) {

								case MARKUP:
									styleRange.start = xr.getStart();
									styleRange.length = regionLength;
									styleRange.fontStyle = SWT.BOLD;
									styleRange.foreground = display.getSystemColor(SWT.COLOR_BLUE);
									txtResponse.setStyleRange(styleRange);
									break;
								case ATTRIBUTE:
									styleRange.start = xr.getStart();
									styleRange.length = regionLength;
									// styleRange.fontStyle = SWT.BOLD;
									styleRange.foreground = display.getSystemColor(SWT.COLOR_DARK_GREEN);
									txtResponse.setStyleRange(styleRange);
									break;
								case ATTRIBUTE_VALUE:
									styleRange.start = xr.getStart();
									styleRange.length = regionLength;
									// styleRange.fontStyle = SWT.BOLD;
									styleRange.foreground = display.getSystemColor(SWT.COLOR_BLACK);
									txtResponse.setStyleRange(styleRange);
									break;
								case MARKUP_VALUE:
									styleRange.start = xr.getStart();
									styleRange.length = regionLength;
									// styleRange.fontStyle = SWT.BOLD;
									styleRange.foreground = display.getSystemColor(SWT.COLOR_RED);
									txtResponse.setStyleRange(styleRange);
									break;
								case COMMENT:
									break;
								case INSTRUCTION:
									break;
								case CDATA:
									break;
								case WHITESPACE:
									break;
								default:
									break;
								}

							}
						} catch (Exception e) {
							txtResponse.setText("Error Occured.");
						}
					}
				});
			}
		});

	}

	private void CleanUIResources() {
		if (btnChooseMethod != null && !btnChooseMethod.isDisposed()) {
			btnChooseMethod.dispose();
		}
		if (btnCreateExcel != null && !btnCreateExcel.isDisposed()) {
			btnCreateExcel.dispose();
		}
		if (btnCreateXml != null && !btnCreateXml.isDisposed()) {
			btnCreateXml.dispose();
		}
		if (btnHitService != null && !btnHitService.isDisposed()) {
			btnHitService.dispose();
		}
		if (txtWsdlUrl != null && !txtWsdlUrl.isDisposed()) {
			txtWsdlUrl.dispose();
		}
		if (listMethods != null && !listMethods.isDisposed()) {
			listMethods.dispose();
		}
		if (listXMLFiles != null && !listXMLFiles.isDisposed()) {
			listXMLFiles.dispose();
		}
		if (txtResponse != null && !txtResponse.isDisposed()) {
			txtResponse.dispose();
		}
		if (display != null && !display.isDisposed()) {
			display.dispose();
		}
		if (shell != null && !shell.isDisposed()) {
			shell.dispose();
		}
	}
}