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

    QDir workDir = QFileInfo(strDocxFileNew).dir();
    QString tmpDir = "_TEMP_";
    workDir.mkdir(tmpDir);
    QString tmpPath = QFileInfo(strDocxFileNew).absolutePath() + QDir::separator() + tmpDir;

    QZipReader zR(&fDocx);
    zR.extractAll(tmpPath);

    QString strrels = tmpPath + QDir::separator() + "_rels/.rels";
    QString strdocPropsapp = tmpPath + QDir::separator() + "docProps/app.xml";
    QString strdocPropscore = tmpPath + QDir::separator() + "docProps/core.xml";
    QString strdocPropscustom = tmpPath + QDir::separator() + "docProps/custom.xml";
    QString strwordrels = tmpPath + QDir::separator() + "word/_rels/document.xml.rels";
    QString strwordtheme = tmpPath + QDir::separator() + "word/theme/theme1.xml";
    QString strworddocument = tmpPath + QDir::separator() + "word/document.xml";
    QString strwordfontTable = tmpPath + QDir::separator() + "word/fontTable.xml";
    QString strwordsettings = tmpPath + QDir::separator() + "word/settings.xml";
    QString strwordstyles = tmpPath + QDir::separator() + "word/styles.xml";
    QString strContentTypes = tmpPath + QDir::separator() + "[Content_Types].xml";

    QFile frels(strrels);
    frels.open(QFile::ReadOnly);

    QFile fdocPropsapp(strdocPropsapp);
    fdocPropsapp.open(QFile::ReadOnly);

    QFile fdocPropscore(strdocPropscore);
    fdocPropscore.open(QFile::ReadOnly);

    QFile fdocPropscustom(strdocPropscustom);
    fdocPropscustom.open(QFile::ReadOnly);

    QFile fwordrels(strwordrels);
    fwordrels.open(QFile::ReadOnly);

    QFile fwordtheme(strwordtheme);
    fwordtheme.open(QFile::ReadOnly);

    QFile fworddocument(strworddocument);
    fworddocument.open(QFile::ReadOnly);

    QFile fwordfontTable(strwordfontTable);
    fwordfontTable.open(QFile::ReadOnly);

    QFile fwordsettings(strwordsettings);
    fwordsettings.open(QFile::ReadOnly);

    QFile fwordstyles(strwordstyles);
    fwordstyles.open(QFile::ReadOnly);

    QFile fContentTypes(strContentTypes);
    fContentTypes.open(QFile::ReadOnly);



    QDomDocument doc;
    doc.setContent(&fworddocument);

    QDomElement root = doc.documentElement();

    QDomNodeList nList = root.elementsByTagName("w:t");
    for (int i = 0 ; i < nList.size(); i++)
    {
        QDomElement ele = nList.at(i).toElement();
        QDomNode n = ele.firstChild();

        QString strText = n.nodeValue();
        QString strTextNew = strText.trimmed();

        strTextNew.prepend("　　");

        ele.firstChild().setNodeValue(strTextNew);

        QDomNode nNew = ele.firstChild();
        ele.replaceChild(nNew, n);
    }

    // <w:ind w:firstLine="420" />
    // to <w:ind w:firstLine="0" />
    QDomNodeList wList = root.elementsByTagName("w:ind");
    for (int i = 0 ; i < wList.size(); i++)
    {
        QDomElement ele = wList.at(i).toElement();
        QDomNode n = ele.firstChild();

        QDomNode nNew = ele.firstChild();
        ele.attributeNode("w:firstLine").setValue("0");


        ele.replaceChild(nNew, n);
    }



    QZipWriter zW(&fDocxNew);

    zW.addDirectory("_rels");
    zW.addDirectory("docProps");
    zW.addDirectory("word");
    zW.addDirectory("word/_rels/");
    zW.addDirectory("word/theme");

    zW.addFile("_rels/.rels", frels.readAll());
    zW.addFile("docProps/app.xml", fdocPropsapp.readAll());
    zW.addFile("docProps/core.xml", fdocPropscore.readAll());
    zW.addFile("docProps/custom.xml", fdocPropscustom.readAll());
    zW.addFile("word/_rels/document.xml.rels", fwordrels.readAll());
    zW.addFile("word/theme/theme1.xml", fwordtheme.readAll());
    zW.addFile("word/document.xml", doc.toByteArray());
    zW.addFile("word/fontTable.xml", fwordfontTable.readAll());
    zW.addFile("word/settings.xml", fwordsettings.readAll());
    zW.addFile("word/styles.xml", fwordstyles.readAll());
    zW.addFile("[Content_Types].xml", fContentTypes.readAll());
    zW.close();

    frels.close();
    fdocPropsapp.close();
    fdocPropscore.close();
    fdocPropscustom.close();
    fwordrels.close();
    fwordtheme.close();
    fworddocument.close();
    fwordfontTable.close();
    fwordsettings.close();
    fwordstyles.close();
    fContentTypes.close();

    zR.close();

    QDir tmpPathDir(tmpPath);
    tmpPathDir.removeRecursively();

    return true;

}
