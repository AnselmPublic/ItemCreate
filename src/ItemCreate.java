
/**********************************************
 * [Copyright ©]
 * @File: ItemCreate.java 
 * @author Frank-Wang
 * @Date: 2019.04.26
 * @Version: 1.1
 * @Since: JDK 1.8.0_92, AGILE 9.3.6
 **********************************************/
import java.io.IOException;
import java.util.ArrayList;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.IDataObject;
import com.agile.api.IItem;
import com.agile.api.IManufacturingSite;
import com.agile.api.INode;
import com.agile.api.ITable;
import com.agile.api.ItemConstants;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;
import com.agile.px.ICreateEventInfo;
import com.agile.px.IEventAction;
import com.agile.px.IEventInfo;

import com.anselm.tools.agile.AUtil;
import com.anselm.tools.record.*;

/**
 * &emsp;&emsp;&emsp;&emsp;<em><b>Class ItemCreate</em></b><BR/>
 * <BR/>
 * This class offer a program to put all the sites in the new create item's site
 * table,and can be controlled by excel to join the sites.
 * 
 * @Step 1. Initialize log file.
 * @Step 2. Get admin session.
 * @Step 3. Get the new create item.
 * @Step 4. Read excel to determine which sites to join.
 * @Step 5. Convert acquired data into IdataObject.
 * @Step 6. Get item's site table.
 * @Step 7. Add all sites in the new create item's site table.
 */
public class ItemCreate implements IEventAction {
	/**
	 * &emsp;&emsp;&emsp;&emsp;<em><b>doAction</em></b><BR/>
	 * <BR/>
	 * &emsp;<font size=
	 * "1">{@linkplain EventActionResult}&emsp;doAction({@linkplain IAgileSession}
	 * session, {@linkplain INode} actionNode, {@linkplain IEventInfo}
	 * req)</font><BR/>
	 * <BR/>
	 * Main function of the class.
	 * 
	 * @param session    {@link com.agile.api.IAgileSession}
	 * @param actionNode {@link com.agile.api.INode}
	 * @param req        {@link com.agile.px.IEventInfo}
	 * 
	 * @return Showing user message.For detailed visits, please refer to
	 *         {@link com.agile.px.EventActionResult}.
	 * @exception APIException
	 * @exception Exception
	 */
	public EventActionResult doAction(IAgileSession session, INode actionNode, IEventInfo req) {
		/** It is used for config reading and writing. */
		Ini ini = null;
		Log log = null;
		IAgileSession admin = null;

		try {
			// Step 1. Initialize log file
			ini = new Ini();
			log = new Log();

			/* Record log path */
			String strLogfilepath = ini.getValue("Program Use", "LogFile");

			log.setLogFile(strLogfilepath + "/SN12_ItemCreate.log", ini);
			log.logSeparatorBar();
			log.setMaxLengthOfName(13);
			// Step 2. Get admin session
			log.log("▶ Get Admin session ..");
			admin = getAdminSession(ini);

			// Step 3. Get the new create item
			ICreateEventInfo icefInfo = (ICreateEventInfo) req;
			IDataObject idoItem = icefInfo.getDataObject();
			idoItem = (IDataObject) admin.getObject(IItem.OBJECT_TYPE, idoItem.getName());
			String strItemAPIName = idoItem.getAgileClass().getAPIName().toLowerCase().trim();

			// Step 4. Read excel to determine which sites to join
			ArrayList<String> arylistData = ReadWriteExcel.ReadExcelFile(ini.getValue("ItemCreate", "EXCEL_FILE_PATH"),
					"廠區", strItemAPIName, log);
			log.logger(log.getArrayToString(arylistData));

			// Step 5. Convert acquired data into IdataObject
			ArrayList<IDataObject> arylistAllSites = getAllSites(admin, arylistData, log);

			// Step 6. Get item's site table
			log.log("▶ Get Site Table ..");
			ITable itableSite = idoItem.getTable(ItemConstants.TABLE_SITES);
			log.log(1, "Item: " + idoItem.getName());

			// Step 7. Add all sites in the new create item's site table
			log.log("▶ Add " + arylistAllSites.size() + " results in Site Table ..");
			for (int i = 0; i < arylistAllSites.size(); i++) {
				itableSite.createRow(arylistAllSites.get(i));
			}

			return new EventActionResult(req,
					new ActionResult(ActionResult.STRING, "Done Adding " + arylistAllSites.size() + " in Site Table."));
		} catch (APIException apie) {
			log.logException(apie);
			apie.printStackTrace();
			return new EventActionResult(req, new ActionResult(ActionResult.EXCEPTION, apie));
		} catch (Exception e) {
			log.logException(e);
			e.printStackTrace();
			return new EventActionResult(req, new ActionResult(ActionResult.EXCEPTION, e));
		} finally {
			close(ini, log);
			admin.close();
		}
	}

	/**
	 * &emsp;&emsp;&emsp;&emsp;<em><b>close</em></b><BR/>
	 * &emsp;<font size="1">close({@linkplain Ini} ini, {@linkplain LogNew}
	 * log)</font><BR/>
	 * <BR/>
	 * Close log and ini construction.
	 * 
	 * @param ini All tools for config.For detailed visits, please refer to
	 *            {@link util.Ini}.
	 * @param log All tools for log.For detailed visits, please refer to
	 *            {@link record.LogNew}.
	 * 
	 * @return No return.
	 * @exception IOException
	 */
	private void close(Ini ini, Log log) {
		log.log("▶Close..");

		ini = null;
		try {
			log.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * &emsp;&emsp;&emsp;&emsp;<em><b>getAdminSession</em></b><BR/>
	 * <BR/>
	 * &emsp;<font size=
	 * "1">{@linkplain IAgileSession}&emsp;getAdminSession({@linkplain Ini}
	 * ini)</font><BR/>
	 * <BR/>
	 * Obtain system administrator authority information by reading config. For
	 * detailed visits, please refer to {@link util.AUtil}.
	 * 
	 * @param ini All tools for config.For detailed visits, please refer to
	 *            {@link util.Ini}.
	 * 
	 * @return Administrator session.For detailed visits, please refer to
	 *         {@link com.agile.api.IAgileSession}.
	 */
	private IAgileSession getAdminSession(Ini ini) {
		String strAgileUrl = ini.getValue("AgileAP", "url");
		String strAgileAdmin = ini.getValue("AgileAP", "username");
		String strAgilepwd = ini.getValue("AgileAP", "password");
		IAgileSession admin = AUtil.getAgileSession(strAgileUrl, strAgileAdmin, strAgilepwd);

		return admin;
	}

	/**
	 * &emsp;&emsp;&emsp;&emsp;<em><b>getAllSites</em></b><BR/>
	 * <BR/>
	 * &emsp;<font size=
	 * "1">{@linkplain ArrayList}&emsp;getAllSites({@linkplain IAgileSession} session)
	 * throws {@linkplain APIException}</font><BR/>
	 * <BR/>
	 * Turning sites data from type {@linkplain String} to {@linkplain IDataObject}.
	 * 
	 * @param session For detailed visits, please refer to
	 *                {@link com.agile.api.IAgileSession}
	 * @param arydata All sites in type of String.
	 * @param log     All tools for log.For detailed visits, please refer to
	 *                {@link record.LogNew}.
	 * 
	 * @return Results of all sites.
	 * @throws APIException
	 * @throws NullPointerException
	 */
	private ArrayList<IDataObject> getAllSites(IAgileSession admin, ArrayList<String> arydata, Log log) {
		ArrayList<IDataObject> arylistAllSites = new ArrayList<>();
		try {
			log.log("▶Turn string data to IDataObject data..");
			for (int i = 0; i < arydata.size(); i++) {
				IDataObject iDataObject = (IDataObject) admin.getObject(IManufacturingSite.OBJECT_TYPE, arydata.get(i));
				arylistAllSites.add(iDataObject);
			}
		} catch (APIException e) {
			e.printStackTrace();
			log.logException(e);
		} catch (NullPointerException e) {
			log.log("Please check excel data. Sites ends with [end].");
			e.printStackTrace();
			log.logException(e);
		}

		return arylistAllSites;
	}
}
