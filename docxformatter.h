#ifndef DOCXFORMATTER_H
#define DOCXFORMATTER_H

#include <QObject>

#include <QFile>
#include <QDir>
#include <QDomDocument>
#include <QXmlStreamWriter>
#include <QDebug>

#include <QtGui/private/qzipreader_p.h>
#include <QtGui/private/qzipwriter_p.h>

class DocxFormatter : public QObject
{
    Q_OBJECT
public:
    explicit DocxFormatter(QObject *parent = nullptr);

    bool process(QString strDocxFile, QString strDocxFileNew, QString &strErrMsg);

signals:

};

#endif // DOCXFORMATTER_H
