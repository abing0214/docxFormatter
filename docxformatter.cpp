#include "docxformatter.h"

DocxFormatter::DocxFormatter(QObject *parent)
    : QObject{parent}
{

}

bool DocxFormatter::process(QString strDocxFile, QString strDocxFileNew, QString &strErrMsg)
{

    QFile fDocx(strDocxFile);
    bool isOpen = fDocx.open(QFile::ReadOnly);
    if (!isOpen)
    {
        strErrMsg = "打开源文件失败，请关闭所有WORD文档后重试。";
        return false;
    }

    QFile fDocxNew(strDocxFileNew);
    isOpen = fDocxNew.open(QFile::WriteOnly);
    if (!isOpen)
    {
        strErrMsg = "打开新文件失败，请关闭所有WORD文档后重试。";
        return false;
    }

    QZipWriter zipW(&fDocxNew);

    zipW.addDirectory("_rels");
    zipW.addDirectory("docProps");
    zipW.addDirectory("word");
    zipW.addDirectory("word/_rels/");
    zipW.addDirectory("word/theme");

    QZipReader zR(&fDocx);
    QVector<QZipReader::FileInfo> fiList = zR.fileInfoList();
    for (int i = 0; i < fiList.size(); i++)
    {
        QZipReader::FileInfo fi = fiList.at(i);

        // "word/document.xml"
        if (fi.filePath == "word/document.xml")
        {
            QDomDocument doc;
            doc.setContent(zR.fileData(fi.filePath));

            QDomElement root = doc.documentElement();

            QDomNodeList wpList = root.elementsByTagName("w:p");
            {
                for (int j = 0; j < wpList.size(); j++)
                {
                    QDomElement wp = wpList.at(j).toElement();
                    QDomNodeList nList = wp.elementsByTagName("w:t");


                    QString strWPText = "";
                    for (int i = 0 ; i < nList.size(); i++)
                    {
                        QDomElement ele = nList.at(i).toElement();
                        QDomNode n = ele.firstChild();
                        QString strText = "";
                        if (n.isNull())
                        {
                             // <w:t xml:space="preserve"> </w:t>
                            if (ele.attributeNode("xml:space").value() == "preserve")
                            {
                                strText = " ";
                            }
                        }
                        else
                        {
                            strText = n.nodeValue();
                        }
                        strWPText.append(strText);
                    }
                    strWPText = strWPText.trimmed();
                    strWPText.prepend("　　");


                    int first = 0;
                    for (int i = 0 ; i < nList.size(); i++)
                    {
                        QDomElement ele = nList.at(i).toElement();
                        QDomNode n = ele.firstChild();

                        if (n.isNull())
                        {
                            first ++;
                            continue;
                        }

                        if (i == first)
                        {
                            ele.firstChild().setNodeValue(strWPText);
                        }
                        else
                        {
                            ele.firstChild().setNodeValue("");
                        }

                        QDomNode nNew = ele.firstChild();
                        ele.replaceChild(nNew, n);
                    }
                }
            }

            // <w:ind w:firstLineChars="200" w:firstLine="420" />
            // to <w:ind w:firstLineChars="0" w:firstLine="0" />
            QDomNodeList wList = root.elementsByTagName("w:ind");
            for (int i = 0 ; i < wList.size(); i++)
            {
                QDomElement ele = wList.at(i).toElement();
                QDomNode n = ele.firstChild();

                QDomNode nNew = ele.firstChild();
                if (ele.hasAttribute("w:firstLine"))
                {
                    ele.attributeNode("w:firstLine").setValue("0");
                }
                if (ele.hasAttribute("w:firstLineChars"))
                {
                    ele.attributeNode("w:firstLineChars").setValue("0");
                }

                ele.replaceChild(nNew, n);
            }

            zipW.addFile(fi.filePath, doc.toByteArray());
        }
        else
        {
            zipW.addFile(fi.filePath, zR.fileData(fi.filePath));
        }
    }

    zipW.close();
    zR.close();

    return true;

}
