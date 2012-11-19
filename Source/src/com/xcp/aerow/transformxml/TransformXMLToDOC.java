/**
 * =====================================================================
 * Project: Transform XML to Office Word Document
 *
 * History (only for major revisions):
 * Date         Author         						            Reason for revision
 * 2011-05-01   Transform XML to PDF activity templates         Creation
 *              (xCP Activity Templates xCelerator)
 * 2012-10-30   Mushiirah MOHUN     				            Migration to xCP 2.0
 * 
 * @com.emc.xcelerator.activities;
 * =====================================================================
 * Copyright (c) 2012 EMC
 * =====================================================================
 */
package com.xcp.aerow.transformxml;

import java.io.IOException;
import java.util.Calendar;
import java.util.logging.FileHandler;
import java.util.logging.Logger;
import java.util.logging.SimpleFormatter;

import com.documentum.com.DfClientX;
import com.documentum.com.IDfClientX;
import com.documentum.fc.client.DfSingleDocbaseModule;
import com.documentum.fc.client.IDfFolder;
import com.documentum.fc.client.IDfModule;
import com.documentum.fc.client.IDfSession;
import com.documentum.fc.client.IDfSysObject;
import com.documentum.fc.common.DfException;
import com.documentum.fc.common.DfLogger;
import com.documentum.operations.IDfXMLTransformNode;
import com.documentum.operations.IDfXMLTransformOperation;

/**
 * The Class TransformXMLToDOC.
 */
public class TransformXMLToDOC extends DfSingleDocbaseModule implements
    IDfModule
{

  /** The session. */
  private IDfSession session;
  
  /** The logger. */
  private Logger logger;

  /**
   * Transform.
   * @param xml_doc_path
   *          the xml_doc_path path where the xml document is saved in the docbase
   * @param xsl_doc_path
   *          the xsl_doc_path path where the XSL document is saved in the docbase
   * @param newObject
   *          the newObject boolean value which indicates if a new object should be created or if it should be set as
   *          rendition
   * @param newObjType
   *          the newObjType type of document that need to be created in docbase
   * @param newObjLocation
   *          the newObjLocation path where the new object created should be stored
   * @param format
   *          the format of the document to be created (example: .doc)
   * @param attrForName
   *          the attrForName the name to be given to the created document
   * @param addDate
   *          the addDate boolean value which indicates whether to append date to the document object_name
   * @param attrForFolder
   *          the attrForFolder append to base folder (optional)
   * @throws DfException
   */
  public void transform(String xml_doc_path, String xsl_doc_path,
      boolean newObject, String newObjType, String newObjLocation,
      String format, String attrForName, boolean addDate, String attrForFolder)
  {

    try
    {
      
      logger = Logger.getLogger("MyLog");  
      FileHandler fh = new FileHandler("C:/temp/MyLogFile.log");  
      logger.addHandler(fh);  

    SimpleFormatter formatter = new SimpleFormatter();  
     fh.setFormatter(formatter);  
      
     logger.info("Obtaining a session");  
      DfLogger.info(this, "Obtaining a session", null, null);
      session = getSession();

      IDfClientX clientx = new DfClientX();

      logger.info("Retrieving XML document from docbase path: "
          + xml_doc_path); 
      DfLogger.info(this, "Retrieving XML document from docbase path: "
          + xml_doc_path, null, null);
      IDfSysObject xmlDoc = (IDfSysObject) session
          .getObjectByPath(xml_doc_path);

      logger.info("Perform Transformation");
      DfLogger.info(this, "Perform Transformation", null, null);
      IDfXMLTransformOperation transformOperation = clientx
          .getXMLTransformOperation();
      transformOperation.setSession(session);
      IDfXMLTransformNode transformNode = (IDfXMLTransformNode) transformOperation
          .add(xmlDoc);

      logger.info("Retrieving XSL document from docbase path: "
          + xsl_doc_path); 
      DfLogger.info(this, "Retrieving XSL document from docbase path: "
          + xsl_doc_path, null, null);
      IDfSysObject xslObject = (IDfSysObject) session
          .getObjectByPath(xsl_doc_path);

      transformOperation.setTransformation(xslObject);
      if (newObject)
      {
        String xmlDocName = attrForName;
        String newxmlDocName = null;
        if (xmlDocName.endsWith(".xml"))
        {
          int dotIndex = xmlDocName.lastIndexOf('.');
          xmlDocName = xmlDocName.substring(0, dotIndex);
        }
        if (addDate)
        {
          Calendar rightNow = Calendar.getInstance();
          int year = rightNow.get(Calendar.YEAR);
          int month = rightNow.get(Calendar.MONTH) + 1;
          int day = rightNow.get(Calendar.DAY_OF_MONTH);
          xmlDocName = xmlDocName + " " + month + "-" + day + "-" + year;
        }
        if (attrForFolder != null && attrForFolder.length() > 0)
        {
          newObjLocation = newObjLocation + "/" + attrForFolder;
          // Create new folder if it doesn't exist
          IDfFolder targetFolderObj = session.getFolderByPath(newObjLocation);
          if (targetFolderObj == null)
          {
            createFolderStructure(session, newObjLocation);
          }
        }
        newxmlDocName = xmlDocName + "." + format;

        if (format.equalsIgnoreCase("doc"))
        {
          format = "msw8";
        }

        logger.info("Creating the new Office Word Document in the docbase path: "
            + newObjLocation); 
        DfLogger.info(this,
            "Creating the new Office Word Document in the docbase path: "
                + newObjLocation, null, null);

        IDfSysObject theNewObject = (IDfSysObject) session
            .newObject(newObjType);
        theNewObject.setContentType(format);
        
        logger.info("linking doc to folder: " + newObjLocation); 
        DfLogger.info(this, "linking doc to folder: " + newObjLocation, null, null);
        theNewObject.link(newObjLocation);
        theNewObject.setObjectName(newxmlDocName);
        theNewObject.save();

        transformOperation.setDestination(theNewObject);
        transformNode.setOutputFormat(format);
        boolean flag = transformOperation.execute();
        theNewObject.save();
        
        logger.info("doc id: " + theNewObject.getObjectId()); 
        DfLogger.info(this, "doc id: " + theNewObject.getObjectId(), null, null);
        
        logger.info("object name: " + theNewObject.getObjectName()); 
        DfLogger.info(this, "object name: " + theNewObject.getObjectName(), null, null);
        
        logger.info("creation date: " + theNewObject.getCreationDate()); 
        DfLogger.info(this, "creation date: " + theNewObject.getCreationDate(), null, null);

      }
      else
      {
        if (format.equalsIgnoreCase("doc"))
        {
          format = "msw8";
        }
        transformNode.setOutputFormat(format);
        boolean flag = transformOperation.execute();
      }
    }
    catch (DfException e)
    {
      logger.info("DfException: " + e); 
      DfLogger.error(this, "DfException: ", null, e);
    }
    catch (SecurityException e)
    {
      logger.info("SecurityException: " + e); 
      DfLogger.error(this, "SecurityException: ", null, e);
    }
    catch (IOException e)
    {
      logger.info("IOException: " + e); 
      DfLogger.error(this, "IOException: ", null, e);
    }
    finally
    {
      DfLogger.info(this, "Releasing Session", null, null);
      session.getSessionManager().release(session);
    }
  }

  /**
   * Creates the folder structure from a given path.
   * @param session
   *          the session instance of IDfSession
   * @param path
   *          the path corresponding to the folder structure to be created
   * @throws DfException
   *           the DfException Signals that an DfException has occurred.
   */
  protected void createFolderStructure(IDfSession session, String path)
      throws DfException
  {
    /*
     * This creates a folder structure from a path (i.e.
     * /Temp/folder1/folder2). It assumes that at least the cabinet already
     * exists. It goes through each level and creates and links a new folder
     * if it doesn't exist.
     */
    String currentPath = null;
    String previousPath = null;
    IDfFolder currentPathObj;

    String[] folderNames = path.split("/");
    currentPath = "/" + folderNames[1];
    if (folderNames.length > 2)
    {
      for (int i = 2; i < folderNames.length; i++)
      {
        previousPath = currentPath;
        currentPath = currentPath + "/" + folderNames[i];
        currentPathObj = session.getFolderByPath(currentPath);
        if (currentPathObj == null)
        {
          currentPathObj = (IDfFolder) session.newObject("dm_folder");
          currentPathObj.setObjectName(folderNames[i]);
          currentPathObj.link(previousPath);
          currentPathObj.save();
        }
      }
    }
  }
}
