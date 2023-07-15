#include "mainwindow.h"
#include "ui_mainwindow.h"

#include "docxformatter.h"

#include <QMessageBox>
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    setWindowTitle("WORD格式小工具");

    mStrDocxFile = "";
    mStrDocxFileNew = "";
}

MainWindow::~MainWindow()
{
    delete ui;
}


void MainWindow::on_bt_select_clicked()
{
    QString file = QFileDialog::getOpenFileName(this, "WORD源文件", "", "*.docx");

    if (!file.isEmpty())
    {
        mStrDocxFile = file;
        mStrDocxFileNew = mStrDocxFile;
        mStrDocxFileNew.replace(".docx", "_NEW.docx");

        ui->le_docx->setText(mStrDocxFile);
        ui->le_docxnew->setText(mStrDocxFileNew);
    }
}


void MainWindow::on_bt_start_clicked()
{
    if (mStrDocxFile.isEmpty())
    {
        QMessageBox::information(this, "错误", "WORD源文件不能为空");
        return;
    }


    QString err;
    DocxFormatter f;
    bool r = f.process(mStrDocxFile, mStrDocxFileNew, err);
    if (r)
    {
        QString strMsg = "处理完成，请查看新文件：\n" + mStrDocxFileNew;
        ui->textEdit->setText(strMsg);
    }
    else
    {
        ui->textEdit->setText(err);
    }
}

