#include "PDFCreator.h"
#include<QDebug>
#include"PythonOffice.h"


PDFCreator::PDFCreator(QWidget *parent): QWidget(parent)
{
    this->InitObjects();
    this->InitGUI();
    this->connectSignalAndSlot();

    // 加载python模块
    QString error_code;
}

PDFCreator::~PDFCreator()
{
    // (Q.*\* )(.*) = NULL;
    // this->$2->deleteLater();

    this->vBox->deleteLater();
    this->foldPathGroupBox->deleteLater();
    this->foldPathVBoxLayout->deleteLater();
    // 选择源文件夹
    this->sourceHBoxLayout->deleteLater();
    this->sourceTipLabel->deleteLater();
    this->sourcePathLineEdit->deleteLater();
    this->sourceSelectFoldPushButton->deleteLater();
    this->sourceOpenPushButton->deleteLater();
    // 选择目标文件夹
    this->destinationHBoxLayout->deleteLater();
    this->destinationTipLabel->deleteLater();
    this->destinationPathLineEdit->deleteLater();
    this->destinationSelectFoldPushButton->deleteLater();
    this->destinationPathAutoComputePushButton->deleteLater();
    this->destinationOpenPushButton->deleteLater();

    // 文件类型
    this->fileTypeGroupBox->deleteLater();
    this->fileTypeHBox->deleteLater();
    this->pptCheckBox->deleteLater();
    this->excelCheckBox->deleteLater();
    this->wordCheckBox->deleteLater();
    this->pictureCheckBox->deleteLater();

    this->outputSettingHBoxLayout->deleteLater();

    // 子文件
    this->subfoldersGroupBox->deleteLater();
    this->subfoldersVBoxLayout->deleteLater();
    this->subfoldersButtonGroup->deleteLater();
    this->enableSubfoldersRadioButton->deleteLater();
    this->disableSubfoldersRadioButton->deleteLater();

    // 文件夹结构
    this->foldStructureGroupBox->deleteLater();
    this->foldStructureVBoxLayout->deleteLater();
    this->foldStructureButtonGroup->deleteLater();
    this->enableFoldStructureRadioButton->deleteLater();
    this->disableFoldStructureRadioButton->deleteLater();

    // 拷贝无关文件
    this->ignoreGroupBox->deleteLater();
    this->ignoreVBoxLayout->deleteLater();
    this->ignoreButtonGroup->deleteLater();
    this->enableIgnoreRadioButton->deleteLater();
    this->disableIgnoreRadioButton->deleteLater();

    // 输出
    this->outputGroupBox->deleteLater();
    this->outputHBoxLayout->deleteLater();
    this->outputTextEdit->deleteLater();

    // 程序控制
    this->runProgressBar->deleteLater();
    this->runProgressTipLabel->deleteLater();
    this->runPushButton->deleteLater();

    this->runControlHBoxLayout->deleteLater();

}

// 初始化GUI
void PDFCreator::InitGUI()
{
    // 文件夹设置
    // 源文件夹
    this->sourceHBoxLayout->addWidget(this->sourceTipLabel,0);
    this->sourceHBoxLayout->addWidget(this->sourcePathLineEdit,1);
    this->sourceHBoxLayout->addWidget(this->sourceSelectFoldPushButton,0); 
    this->sourceHBoxLayout->addWidget(this->sourceOpenPushButton, 0);
    this->sourceTipLabel->setMinimumWidth(90);
    // 目标文件夹   
    this->destinationHBoxLayout->addWidget(this->destinationTipLabel,0);
    this->destinationHBoxLayout->addWidget(this->destinationPathLineEdit,1);
    this->destinationHBoxLayout->addWidget(this->destinationPathAutoComputePushButton,0);
    this->destinationHBoxLayout->addWidget(this->destinationSelectFoldPushButton,0);
    this->destinationHBoxLayout->addWidget(this->destinationOpenPushButton, 0);
    this->destinationTipLabel->setMinimumWidth(90);
    // 文件夹Groupbox
    this->foldPathVBoxLayout->addLayout(this->sourceHBoxLayout);
    this->foldPathVBoxLayout->addLayout(this->destinationHBoxLayout);
    this->foldPathGroupBox->setLayout(this->foldPathVBoxLayout);

    // 文件类型
    this->pptCheckBox->setChecked(true);
    this->excelCheckBox->setChecked(false);
    this->wordCheckBox->setChecked(true);
    this->pictureCheckBox->setChecked(true);
    this->fileTypeHBox->addWidget(this->pptCheckBox);
    //this->fileTypeHBox->addWidget(this->excelCheckBox);
    this->fileTypeHBox->addWidget(this->wordCheckBox);
    this->fileTypeHBox->addWidget(this->pictureCheckBox);
    this->fileTypeGroupBox->setLayout(this->fileTypeHBox);

    // 子文件
    this->subfoldersButtonGroup->addButton(this->enableSubfoldersRadioButton);
    this->subfoldersButtonGroup->addButton(this->disableSubfoldersRadioButton);
    this->subfoldersVBoxLayout->addWidget(this->enableSubfoldersRadioButton);
    this->subfoldersVBoxLayout->addWidget(this->disableSubfoldersRadioButton);
    this->subfoldersGroupBox->setLayout(this->subfoldersVBoxLayout);

    // 文件夹结构
    this->foldStructureButtonGroup->addButton(this->enableFoldStructureRadioButton);
    this->foldStructureButtonGroup->addButton(this->disableFoldStructureRadioButton);
    this->foldStructureVBoxLayout->addWidget(this->enableFoldStructureRadioButton);
    this->foldStructureVBoxLayout->addWidget(this->disableFoldStructureRadioButton);
    this->foldStructureGroupBox->setLayout(this->foldStructureVBoxLayout);

    // 拷贝无关文件
    this->ignoreButtonGroup->addButton(this->enableIgnoreRadioButton);
    this->ignoreButtonGroup->addButton(this->disableIgnoreRadioButton);
    this->ignoreVBoxLayout->addWidget(this->enableIgnoreRadioButton);
    this->ignoreVBoxLayout->addWidget(this->disableIgnoreRadioButton);
    this->ignoreGroupBox->setLayout(this->ignoreVBoxLayout);

    // 输出设置
    this->outputSettingHBoxLayout->addWidget(this->subfoldersGroupBox);
    this->outputSettingHBoxLayout->addWidget(this->foldStructureGroupBox);
    this->outputSettingHBoxLayout->addWidget(this->ignoreGroupBox);

    // 程序输出
    this->outputTextEdit->setLineWrapMode(QTextEdit::NoWrap);
    this->outputHBoxLayout->addWidget(this->outputTextEdit);
    this->outputGroupBox->setLayout(this->outputHBoxLayout);

    // 程序控制
    this->runControlHBoxLayout->addWidget(this->runProgressTipLabel);
    this->runControlHBoxLayout->addWidget(this->runProgressBar);
    //this->runControlHBoxLayout->addStretch();
    this->runControlHBoxLayout->addWidget(this->runPushButton);

    // 总体布局
    this->vBox->addWidget(this->foldPathGroupBox);
    this->vBox->addWidget(this->fileTypeGroupBox);
    this->vBox->addLayout(this->outputSettingHBoxLayout);
    this->vBox->addWidget(this->outputGroupBox);
    this->vBox->addLayout(this->runControlHBoxLayout);
    this->vBox->setSpacing(10);
    this->setLayout(this->vBox);

    this->outputLine("欢迎：", "本程序由a_flying_fish独立开发，完全免费，如果您使用中遇到问题，请联系：a_flying_fish@outlook.com",Qt::gray);
    this->outputLine("提示：", "office文件转化功能仅在安装了office 2007及以上版本的系统上受支持!\r\n", Qt::blue);
    this->resize(850, 630);
}

// 实例化对象
void PDFCreator::InitObjects()
{
    // (Q.*)\* (.*) = NULL;
    // this->$2 = new $1();

    this->vBox =  new QVBoxLayout();
    this->foldPathGroupBox =  new QGroupBox(tr("文件夹设置:"));
    this->foldPathVBoxLayout =  new QVBoxLayout();
    // 选择源文件夹
    this->sourceHBoxLayout =  new QHBoxLayout();
    this->sourceTipLabel =  new QLabel(tr("源文件夹:"));
    this->sourcePathLineEdit =  new QLineEdit();
    this->sourcePathLineEdit->setPlaceholderText(tr("请输入源文件所在文件夹"));
    this->sourceSelectFoldPushButton =  new QPushButton(tr("选择"));
    this->sourceOpenPushButton = new QPushButton(tr("打开"));
    // 选择目标文件夹
    this->destinationHBoxLayout =  new QHBoxLayout();
    this->destinationTipLabel =  new QLabel(tr("目标文件夹:"));
    this->destinationPathLineEdit =  new QLineEdit();
    this->destinationPathLineEdit->setPlaceholderText(tr("请给出输出文件存储文件夹"));
    this->destinationPathAutoComputePushButton =  new QPushButton(tr("自动"));
    this->destinationPathAutoComputePushButton->setToolTip(tr("自动填充目标文件夹路径"));
    this->destinationSelectFoldPushButton = new QPushButton(tr("选择"));
    this->destinationOpenPushButton = new QPushButton(tr("打开"));

    // 文件类型
    this->fileTypeGroupBox = new QGroupBox(tr("待转化文件类型:"));
    this->fileTypeHBox = new QHBoxLayout();
    this->pptCheckBox = new QCheckBox(tr("*.ppt|*.pptx"));
    this->excelCheckBox = new QCheckBox(tr("*.xls|*.xlsx"));
    this->wordCheckBox = new QCheckBox(tr("*.doc|*.docx"));
    this->pictureCheckBox = new QCheckBox(tr("*.png|*.jpg|*.jpeg|*.bmp"));

    this->outputSettingHBoxLayout = new QHBoxLayout();

    // 子文件
    this->subfoldersGroupBox = new QGroupBox(tr("子文件夹"));
    this->subfoldersVBoxLayout = new QVBoxLayout();
    this->subfoldersButtonGroup = new QButtonGroup();
    this->enableSubfoldersRadioButton = new QRadioButton(tr("考虑在内"));
    this->enableSubfoldersRadioButton->setChecked(true);
    this->disableSubfoldersRadioButton = new QRadioButton(tr("忽略"));

    // 文件夹结构
    this->foldStructureGroupBox = new QGroupBox(tr("目标文件结构"));
    this->foldStructureVBoxLayout = new QVBoxLayout();
    this->foldStructureButtonGroup = new QButtonGroup();
    this->enableFoldStructureRadioButton = new QRadioButton(tr("保持原目录结构"));
    this->enableFoldStructureRadioButton->setChecked(true);
    this->disableFoldStructureRadioButton = new QRadioButton(tr("直接平铺"));

    // 拷贝无关文件
    this->ignoreGroupBox = new QGroupBox(tr("非目标文件"));
    this->ignoreVBoxLayout = new QVBoxLayout();
    this->ignoreButtonGroup = new QButtonGroup();
    this->enableIgnoreRadioButton = new QRadioButton(tr("不出现于目标文件夹"));
    this->enableIgnoreRadioButton->setChecked(true);
    this->disableIgnoreRadioButton = new QRadioButton(tr("复制到目标文件夹中"));

    // 输出
    this->outputGroupBox = new QGroupBox(tr("运行提示: "));
    this->outputHBoxLayout = new QHBoxLayout();
    this->outputTextEdit = new QTextEdit();
    this->outputTextEdit->setReadOnly(true);

    // 程序控制
    this->runProgressBar = new QProgressBar();
    this->runProgressBar->setValue(100);
    this->runProgressTipLabel = new QLabel(tr("运行进度: "));
    this->runPushButton = new QPushButton(tr("运行"));

    this->runControlHBoxLayout = new QHBoxLayout();

}

void PDFCreator::connectSignalAndSlot()
{
    connect(this->sourceSelectFoldPushButton, SIGNAL(clicked(bool)), this, SLOT(selectSourceFlod()));
    connect(this->sourceOpenPushButton, SIGNAL(clicked(bool)), this, SLOT(openSourceFold()));
    connect(this->destinationSelectFoldPushButton, SIGNAL(clicked(bool)), this, SLOT(selectDestinationFold()));
    connect(this->destinationPathAutoComputePushButton, SIGNAL(clicked(bool)), this, SLOT(autoComputeDestinationFold()));
    connect(this->destinationOpenPushButton, SIGNAL(clicked(bool)), this, SLOT(openDestinationFold()));
    connect(this->runPushButton, SIGNAL(clicked(bool)), this, SLOT(startDealFiles()));
}

void PDFCreator::selectSourceFlod()
{
    //文件夹路径
    QString path = QFileDialog::getExistingDirectory(this, tr("选择源文件所在文件夹"), QStandardPaths::writableLocation(QStandardPaths::DesktopLocation));

    if (!path.isEmpty())
    {
        this->sourcePathLineEdit->setText(path);
    }
}

void PDFCreator::openSourceFold()
{
    this->sourcePathLineEdit->setText(this->fixErrorPath(this->sourcePathLineEdit->text()));
    if (!QDir(this->sourcePathLineEdit->text()).exists())
    {
        if (QMessageBox::information(this, tr("提示"), tr("源文件夹不存在，是否先为您生成呢？"), QMessageBox::Yes | QMessageBox::No, QMessageBox::No) == QMessageBox::Yes)
        {
            if (this->ensureFoldExitst(this->sourcePathLineEdit->text()))
            {
                QDesktopServices::openUrl(QUrl("file:///" + this->sourcePathLineEdit->text(), QUrl::TolerantMode));

                this->outputLine("尝试打开不存在的源文件夹（现已新建）:", this->sourcePathLineEdit->text(), Qt::blue); 
            }
            else
            {
                this->outputLine("尝试打开不存在的源文件夹，且无法自动判别新建:", this->sourcePathLineEdit->text(), Qt::red);
            }
            return;
        }

        this->outputLine("尝试打开不存在的源文件夹（失败且未新建）:", this->sourcePathLineEdit->text(),Qt::red);
    }
    else
    {
        QDesktopServices::openUrl(QUrl("file:///" + this->sourcePathLineEdit->text(), QUrl::TolerantMode));
    }
}

void PDFCreator::openDestinationFold()
{
    this->destinationPathLineEdit->setText(this->fixErrorPath(this->destinationPathLineEdit->text()));
    if (!QDir(this->destinationPathLineEdit->text()).exists())
    {
        if (QMessageBox::information(this, tr("提示"), tr("目标文件夹暂不存在，但是有可能在程序运行后生成，是否提前为您创建呢？"), QMessageBox::Yes | QMessageBox::No, QMessageBox::No) == QMessageBox::Yes)
        {
            if (this->ensureFoldExitst(this->destinationPathLineEdit->text()))
            {
                QDesktopServices::openUrl(QUrl("file:///" + this->destinationPathLineEdit->text(), QUrl::TolerantMode));

                this->outputLine("尝试打开不存在的目标文件夹(现已新建)->", this->destinationPathLineEdit->text(), Qt::blue);
            }
            else
            {
                this->outputLine("尝试打开不存在的目标文件夹，且无法自动判别新建:", this->sourcePathLineEdit->text(), Qt::red);
            }
            return;
        }

        this->outputLine("尝试打开不存在的目标文件夹(失败且未新建)->", this->destinationPathLineEdit->text(),Qt::red);
    }
    else
    {
        QDesktopServices::openUrl(QUrl("file:///" + this->destinationPathLineEdit->text(), QUrl::TolerantMode));
    }
}

// 递归运算，确保路径存在
bool PDFCreator::ensureFoldExitst(QString path)
{
    if (!QDir(path).exists())
    {
        // 确保父级路径存在
        QString fatherPath = this->getFatherPath(path);
        if (fatherPath.isEmpty())
        {
            return false;
        }
        QString foldName = this->getFileName(path);
        if (!this->ensureFoldExitst(fatherPath))
        {
            return false;
        }

        QDir(fatherPath).mkdir(foldName);

        // 输出提示
        this->outputLine("创建文件夹：", QDir(fatherPath).absoluteFilePath(foldName));
    }

    return true;
}

void PDFCreator::outputLine(QString tip, QString contain, int color)
{
    this->lock.lock();
    this->outputTextEdit->setTextColor(QColor((Qt::GlobalColor)color));
    this->outputTextEdit->insertPlainText(tip + contain+"\r\n");
    this->outputTextEdit->setTextColor(QColor(Qt::black));
    this->lock.unlock();
}

void PDFCreator::startDealFiles()
{
    if (!this->isRun)
    {
        this->isRun = true;
        // 确保路径合法
        this->sourcePathLineEdit->setText(this->fixErrorPath(this->sourcePathLineEdit->text()));
        this->destinationPathLineEdit->setText(this->fixErrorPath(this->destinationPathLineEdit->text()));

        // 源文件夹检查
        if (this->sourcePathLineEdit->text().size()<1 || !QDir(this->sourcePathLineEdit->text()).exists() || QDir(this->sourcePathLineEdit->text()).entryInfoList().count() <= 2)
        {
            QMessageBox::information(this, tr("提示"), tr("源文件夹不存在，或为空文件夹"));
            this->runPushButton->setText("运行");
            this->runProgressBar->setValue(100);
            this->isRun = false;
            this->isPause = false;
            return;
        }

        // 目标文件夹检查
        if (this->destinationPathLineEdit->text().size() < 1)
        {
            QMessageBox::information(this, tr("提示"), tr("目标存储文件夹路径不能为空"));
            this->runPushButton->setText("运行");
            this->runProgressBar->setValue(100);
            this->isRun = false;
            this->isPause = false;
            return;
        }

        // qDebug() << this->sourcePathLineEdit->text() << QDir(this->sourcePathLineEdit->text()).exists() << QDir(this->sourcePathLineEdit->text()).entryInfoList() << endl;

        // 创建目标文件夹
        if (!this->ensureFoldExitst(this->destinationPathLineEdit->text()))
        {
            QMessageBox::information(this, tr("错误"), tr("目标文件夹路径严重错误，无法自动创建"));
            this->runPushButton->setText("运行");
            this->runProgressBar->setValue(100);
            this->isRun = false;
            this->isPause = false;
            return;
        }
        

        this->runPushButton->setText(tr("暂停"));
        this->isPause = false;
        this->runProgressBar->setValue(0);
        thread = QtConcurrent::run(this, &PDFCreator::runProcess);
        return;
    }
    else
    {
        if (this->isPause)
        {
            this->thread.setPaused(false);
            this->isPause = false;
            this->runPushButton->setText(tr("暂停"));
        }
        else
        {
            this->thread.setPaused(true);
            this->runPushButton->setText(tr("运行"));
            this->isPause = true;
            return;
        }
    }
}

void PDFCreator::flushGUIProgressBar(int value)
{
    this->runProgressBar->setValue(value);
}

void PDFCreator::runProcess()
{
    connect(this, SIGNAL(flushProgressBar(int)), this, SLOT(flushGUIProgressBar(int)));
    connect(this, SIGNAL(sendOutput(QString, QString, int)), this, SLOT(outputLine(QString, QString, int)));

    this->totalFilesCount = countTotalWorks(this->sourcePathLineEdit->text());
    this->dealFilesCount = 0;

    this->recursionDealFile(this->sourcePathLineEdit->text(),this->destinationPathLineEdit->text());
    
    // 运行结束
    this->runPushButton->setText("运行");
    emit this->flushProgressBar(100);           // 更新进度条
    this->isRun = false;
    this->isPause = false;
    this->dealFilesCount = 0;
    this->totalFilesCount = 0;
    this->outputLine("完成...", "\r\n", Qt::lightGray);

    disconnect(this, SIGNAL(flushProgressBar(int)), this, SLOT(flushGUIProgressBar(int)));
    disconnect(this, SIGNAL(sendOutput(QString, QString, int)), this, SLOT(outputLine(QString, QString, int)));
}


bool PDFCreator::ppt_to_pdf(QString sourceFile, QString destinationFile)
{
    /*if (sourceFile.endsWith(".pptx") || sourceFile.endsWith(".ppt"))
    {
        if (QFile(destinationFile).exists())
        {
            emit sendOutput("文件已经存在，忽略：", destinationFile, Qt::blue);
            return true;
        }

        QString error_code = QString();
        if (PythonOffice::word_to_pdf(sourceFile, destinationFile, error_code))
        {
            return true;
        }
        else
        {
            emit sendOutput("错误提示：", error_code, Qt::red);
            return false;
        }
    }
    return true;*/

    if (sourceFile.endsWith(".pptx") || sourceFile.endsWith(".ppt"))
    {
        if (QFile(destinationFile).exists())
        {
            emit sendOutput("文件已经存在，忽略：", destinationFile, Qt::blue);
            return true;
        }

        // Open parameter
        QAxObject* powerPointAxObj = new QAxObject("Powerpoint.Application", 0);
        if (!powerPointAxObj)
        {
            return false;
        }

        powerPointAxObj->dynamicCall("SetVisible(bool)", false);
        QAxObject* presentations = powerPointAxObj->querySubObject("Presentations");
        QList<QVariant> paramList;
        paramList.push_back(QVariant(QString(sourceFile).replace("/", "\\")));
        paramList.push_back(0);
        paramList.push_back(0);
        paramList.push_back(0);

        QAxObject* presentation = presentations->querySubObject("Open(const QString&,int,int,int)", paramList);
        if (presentation != nullptr)
        {
            paramList.clear();
            paramList.push_back(QString(destinationFile).replace("/", "\\"));
            paramList.push_back(32);
            paramList.push_back(0);
            presentation->dynamicCall("SaveAs(const QString&,int,int)", paramList);
            presentations->dynamicCall("Close()");

            delete presentations;
        }

        powerPointAxObj->dynamicCall("Quit(void)");
    }

    return true;
}

bool PDFCreator::word_to_pdf(QString sourceFile, QString destinationFile)
{
    /*if (sourceFile.endsWith(".doc") || sourceFile.endsWith(".docx"))
    {
        if (QFile(destinationFile).exists())
        {
            emit sendOutput("文件已经存在，忽略：", destinationFile, Qt::blue);
            return true;
        }

        QString error_code = QString();
        if (PythonOffice::word_to_pdf(sourceFile, destinationFile, error_code))
        {
            return true;
        }
        else
        {
            emit sendOutput("错误提示：", error_code, Qt::red);
            return false;
        }
    }
    return true;*/

    if (sourceFile.endsWith(".doc") || sourceFile.endsWith(".docx"))
    {
        if (QFile(destinationFile).exists())
        {
            emit sendOutput("文件已经存在，忽略：", destinationFile, Qt::blue);
            return true;
        }

        QAxObject* pWordApplication = new QAxObject("Word.Application", 0);
        if (!pWordApplication)
        {
            return false;
        }

        QAxObject* pWordDocuments = pWordApplication->querySubObject("Documents");

        QVariant filename(QString(sourceFile).replace("/", "\\"));
        QVariant confirmconversions(false);
        QVariant readonly(true);
        QVariant addtorecentfiles(false);
        QVariant passworddocument("");
        QVariant passwordtemplate("");
        QVariant revert(false);
        //打开
        QAxObject* doc = pWordDocuments->querySubObject("Open(const QVariant&, const QVariant&,const QVariant&, "
            "const QVariant&, const QVariant&, "
            "const QVariant&,const QVariant&)",
            filename,
            confirmconversions,
            readonly,
            addtorecentfiles,
            passworddocument,
            passwordtemplate,
            revert);

        QVariant OutputFileName(QString(destinationFile).replace("/", "\\"));
        QVariant ExportFormat(17);                  //17是pdf
        QVariant OpenAfterExport(false);            //保存后是否自动打开
        //转成pdf
        doc->querySubObject("ExportAsFixedFormat(const QVariant&,const QVariant&,const QVariant&)",
            OutputFileName,
            ExportFormat,
            OpenAfterExport
        );

        //关闭
        doc->dynamicCall("Close(boolean)", false);
        pWordApplication->dynamicCall("Quit(void)");
    }

    return true;
}

bool PDFCreator::excel_to_pdf(QString sourceFile, QString destinationFile)
{
    /*if (sourceFile.endsWith(".xls") || sourceFile.endsWith(".xlsx"))
    {
        if (QFile(destinationFile).exists())
        {
            emit sendOutput("文件已经存在，忽略：", destinationFile, Qt::blue);
            return true;
        }

        QString error_code = QString();
        if (PythonOffice::word_to_pdf(sourceFile, destinationFile, error_code))
        {
            return true;
        }
        else
        {
            emit sendOutput("错误提示：", error_code, Qt::red);
            return false;
        }
    }
    return true;*/

    if (sourceFile.endsWith(".xls") || sourceFile.endsWith(".xlsx"))
    {
        if (QFile(destinationFile).exists())
        {
            emit sendOutput("文件已经存在，忽略：", destinationFile, Qt::blue);
            return true;
        }

        QAxObject* excel = new QAxObject("Excel.Application");
        if (!excel)
        {
            return false;
        }
        excel->dynamicCall("SetVisible (bool Visible)", "false");
        QAxObject* workbooks = excel->querySubObject("Workbooks");

        QVariant filename(sourceFile);
        //打开
        QAxObject* workbook = workbooks->querySubObject("Open(const QString&)", filename);

        //QVariant OutputFileName(destinationFile);
        //QVariant ExportFormat(17);                  //17是pdf
        //QVariant OpenAfterExport(false);            //保存后是否自动打开
        ////转成pdf
        //workbook->querySubObject("ExportAsFixedFormat(const QVariant&,const QVariant&,const QVariant&)",
        //    OutputFileName,
        //    ExportFormat,
        //    OpenAfterExport
        //);

        QAxObject* worksheets = workbook->querySubObject("Sheets");
        QAxObject* pagecfg = worksheets->querySubObject("PageSetup");
        //设置工作表 区域为1页
        pagecfg->setProperty("Zoom", false);
        pagecfg->setProperty("FitToPagesWide", 1);
        pagecfg->setProperty("FitToPagesTall", 1);
        delete pagecfg;

        worksheets->dynamicCall("ExportAsFixedFormat(QVariant, QVariant)",0, destinationFile);


        //关闭
        workbook->dynamicCall("Close(boolean)", false);
        excel->dynamicCall("Quit(void)");
    }

    return true;
}

bool PDFCreator::picture_to_pdf(QString sourceFile, QString destinationFile)
{
    try
    {
        if (sourceFile.endsWith(".png") || sourceFile.endsWith(".jpg") || sourceFile.endsWith(".jpeg") || sourceFile.endsWith(".bmp"))
        {
            if (QFile(destinationFile).exists())
            {
                emit sendOutput("文件已经存在，忽略：", destinationFile, Qt::blue);
                return true;
            }

            QPixmap pixmap(sourceFile);

            //图片生成pdf  
            QPrinter printerPixmap(QPrinter::HighResolution);
            printerPixmap.setPageSize(QPrinter::Custom);                                    //设置纸张大小为A4  
            printerPixmap.setOutputFormat(QPrinter::PdfFormat);                             //设置输出格式为pdf  
            printerPixmap.setOutputFileName(destinationFile);                               //设置输出路径  

            double dHeight = 300;
            double dWidth = (double)pixmap.width() / (double)pixmap.height() * dHeight;
            printerPixmap.setPageSizeMM(QSize(dWidth, dHeight));
            printerPixmap.setPageMargins(0, 0, 0, 0, QPrinter::DevicePixel);

            QPainter painterPixmap;
            painterPixmap.begin(&printerPixmap);
            QRect rect = painterPixmap.viewport();
            float multiple = (double(rect.width()) / pixmap.width());
            float yMultiple = (double(rect.height()) / pixmap.height());
            float fMin = (multiple < yMultiple) ? multiple : yMultiple;
            painterPixmap.scale(fMin, fMin);                                                //将图像(所有要画的东西)在pdf上放大multiple-1倍  
            painterPixmap.drawPixmap(0, 0, pixmap.width(), pixmap.height(), pixmap);        //画图  
            painterPixmap.end();
        }

        return true;
    }
    catch (...)
    {
        return false;
    }
}

// 递归
int PDFCreator::countTotalWorks(QString sourcePath)
{
    int totalFiles = 0;

    QDir dir(sourcePath);
    QStringList fileList = dir.entryList(QDir::Files);
    totalFiles += fileList.size();

    QStringList dirList = dir.entryList(QDir::AllDirs | QDir::NoDotAndDotDot);
    for (int i = 0; i < dirList.size(); i++)
    {
        if (dirList.size() < 1)
        {
            break;
        }

        totalFiles += countTotalWorks(dir.absoluteFilePath(dirList[i]));
    }

    return totalFiles;
}

QString PDFCreator::getRealName(QString name)
{
    if (name.isEmpty())
    {
        return QString();
    }

    int i = name.size() - 1;
    while (i>=0)
    {
        if (name[i] == '.')
        {
            i--;
            break;
        }
        i--;
    }

    // 没有'.'
    if (i < 0)
    {
        return QString(name);
    }
    else
    {
       QString temp = QString(name).mid(0, i+1).trimmed();
       if (temp.size() <= 0)
       {
           return "_";
       }
       else
       {
           return temp;
       }
    }

}

// 递归
void PDFCreator::recursionDealFile(QString sourcePath, QString destinationPath)
{
    QDir dir(sourcePath);
    /*
        // ppt转化
        if (!filters_ppt.isEmpty())
        {
            QStringList fileList = dir.entryList(filters_ppt, QDir::Files);
            for (int i = 0; i < fileList.size(); i++)
            {
                if (fileList.size() < 1)
                {
                    break;
                }

                QString file = fileList[i];

                if (!ppt_to_pdf(dir.absoluteFilePath(file), QDir(destinationPath).absoluteFilePath(file.replace(".pptx", ".pdf").replace(".ppt", ".pdf"))))
                {
                    return; //出错
                }
            }
        }

        // excel文件
        if (!filters_excel.isEmpty())
        {
            QStringList fileList = dir.entryList(filters_excel, QDir::Files);
            for (int i = 0; i < fileList.size(); i++)
            {
                if (fileList.size() < 1)
                {
                    break;
                }

                QString file = fileList[i];

                if (!excel_to_pdf(dir.absoluteFilePath(file), QDir(destinationPath).absoluteFilePath(file.replace(".xlsx", ".pdf").replace(".xls", ".pdf"))))
                {
                    return; //出错
                }
            }
        }

        // word文件
        if (!filters_word.isEmpty())
        {
            QStringList fileList = dir.entryList(filters_word, QDir::Files);
            for (int i = 0; i < fileList.size(); i++)
            {
                if (fileList.size() < 1)
                {
                    break;
                }

                QString file = fileList[i];

                if (!word_to_pdf(dir.absoluteFilePath(file), QDir(destinationPath).absoluteFilePath(file.replace(".docx", ".pdf").replace(".doc", ".pdf"))))
                {
                    return; //出错
                }
            }
        }

        // picture文件
        if (!filters_picture.isEmpty())
        {
            QStringList fileList = dir.entryList(filters_picture, QDir::Files);
            for (int i = 0; i < fileList.size(); i++)
            {
                if (fileList.size() < 1)
                {
                    break;
                }

                QString file = fileList[i];

                if (!picture_to_pdf(dir.absoluteFilePath(file), QDir(destinationPath).absoluteFilePath(file.replace(".jpeg", ".pdf").replace(".jpg", ".pdf").replace(".png", ".pdf").replace(".bmp", ".pdf"))))
                {
                    return; //出错
                }
            }
        }

    */

    // 处理文件
    QStringList fileList = dir.entryList(QDir::Files);
    for (int i = 0; i < fileList.size(); i++)
    { 
        emit this->flushProgressBar(int(dealFilesCount * 100 / totalFilesCount)); // 更新进度条

        if (fileList.size() < 1)
        {
            break;
        }

        QString file = fileList[i];

        //// COM组件中，文件名包含空格时存在错误，此处代码绕过(重命名）
        //QString file = QString(file_real).replace(" ", "__");
        //QFile::rename(dir.absoluteFilePath(file_real), dir.absoluteFilePath(file));

        emit sendOutput("开始处理：", dir.absoluteFilePath(file),Qt::black);
        if ((file.endsWith(".pptx") || file.endsWith(".ppt")) && this->pptCheckBox->isChecked())
        {
            QString aimFile = QDir(destinationPath).absoluteFilePath(this->getRealName(file) + ".pdf");
            if (!ppt_to_pdf(dir.absoluteFilePath(file), aimFile))
            {
                emit sendOutput("转PDF出错：", dir.absoluteFilePath(file),Qt::red);
            }
            else
            {
                emit sendOutput("完成：", aimFile, Qt::black);
            }
        }
        else if ((file.endsWith(".xlsx") || file.endsWith(".xls")) && this->excelCheckBox->isChecked())
        {
            QString aimFile = QDir(destinationPath).absoluteFilePath(this->getRealName(file));
            if (!excel_to_pdf(dir.absoluteFilePath(file), aimFile))
            {
                emit sendOutput("转PDF出错：", dir.absoluteFilePath(file), Qt::red);
                continue;
            }
            else
            {
                emit sendOutput("完成：", aimFile, Qt::black);
            }
        }
        else if ((file.endsWith(".docx") || file.endsWith(".doc")) && this->wordCheckBox->isChecked())
        {
            QString aimFile = QDir(destinationPath).absoluteFilePath(this->getRealName(file) + ".pdf");
            if (!word_to_pdf(dir.absoluteFilePath(file), aimFile))
            {

                emit sendOutput("转PDF出错：", dir.absoluteFilePath(file), Qt::red);
            }
            else
            {
                emit sendOutput("完成：", aimFile, Qt::black);
            }
        }
        else if ((file.endsWith(".jpeg") || file.endsWith(".jpg") || file.endsWith(".png") || file.endsWith(".bmp")) && this->pictureCheckBox->isChecked())
        {
            QString aimFile = QDir(destinationPath).absoluteFilePath(this->getRealName(file)+".pdf");
            if (!picture_to_pdf(dir.absoluteFilePath(file), aimFile))
            {
                emit sendOutput("转PDF出错：", dir.absoluteFilePath(file), Qt::red);
            }
            else
            {
                emit sendOutput("完成：", aimFile, Qt::black);
            }
        }
        else
        {
            // 要求拷贝无关文件
            if (this->disableIgnoreRadioButton->isChecked())
            {
                if (QFile(QDir(destinationPath).absoluteFilePath(file)).exists())
                {
                    emit sendOutput("文件已经存在，忽略：", dir.absoluteFilePath(file) + "->" + QDir(destinationPath).absoluteFilePath(file), Qt::blue);
                }
                else
                {
                    if (!QFile::copy(dir.absoluteFilePath(file), QDir(destinationPath).absoluteFilePath(file)))
                    {
                        emit sendOutput("拷贝无关文件时出错：", dir.absoluteFilePath(file) + "->" + QDir(destinationPath).absoluteFilePath(file), Qt::red);
                    }
                    else
                    {
                        emit sendOutput("拷贝文件：", dir.absoluteFilePath(file) + "->" + QDir(destinationPath).absoluteFilePath(file), Qt::black);
                    }
                }
            }
            else
            {
                emit sendOutput("正常忽略...", "", Qt::black);
            }
        }

        /*QFile::rename(dir.absoluteFilePath(file), dir.absoluteFilePath(file_real));
        qDebug() << dir.absoluteFilePath(file) << "->" << dir.absoluteFilePath(file_real) << endl;*/

        emit sendOutput("", "", Qt::black);

        this->dealFilesCount++;
    }

    // 处理文件夹
    if (this->enableSubfoldersRadioButton->isChecked())
    {
        QStringList dirList = dir.entryList(QDir::AllDirs | QDir::NoDotAndDotDot);
        for (int i = 0; i < dirList.size(); i++)
        {
            if (dirList.size() < 1)
            {
                break;
            }

            QString tmp_dir = dirList[i];

            if (this->enableFoldStructureRadioButton->isChecked())
            {
                QString savePath = this->confirmFoldName(QDir(destinationPath).absoluteFilePath(tmp_dir));
                this->ensureFoldExitst(savePath);
                recursionDealFile(dir.absoluteFilePath(tmp_dir), savePath);
            }
            else
            {
                recursionDealFile(dir.absoluteFilePath(tmp_dir), destinationPath);
            }

        }
    }
    return;
}

void PDFCreator::selectDestinationFold()
{
    //文件夹路径
    QString path = QFileDialog::getExistingDirectory(this, tr("选择目标文件存储文件夹"), QStandardPaths::writableLocation(QStandardPaths::DesktopLocation));

    if (!path.isEmpty())
    {
        this->destinationPathLineEdit->setText(path);
    }
}

// 自动计算合适的存储目录
void PDFCreator::autoComputeDestinationFold()
{
    // 源文件夹为空、源文件夹错误，或源文件夹为根目录时，自动指向桌面
    if (this->sourcePathLineEdit->text().isEmpty() || !QDir(this->sourcePathLineEdit->text()).exists() || !QDir(this->sourcePathLineEdit->text()).cdUp())
    {
        this->destinationPathLineEdit->setText(this->confirmFoldName(QDir(QStandardPaths::writableLocation(QStandardPaths::DesktopLocation)).absoluteFilePath("pdf_translate/")));
    }
    else
    {
        QDir sourceDir = QDir(this->sourcePathLineEdit->text());
        QString sourceName = sourceDir.dirName();
        sourceDir.cdUp();
        this->destinationPathLineEdit->setText(this->confirmFoldName(sourceDir.absoluteFilePath(sourceName+"_pdfs/")));
    }
}

QString PDFCreator::confirmFoldName(QString path)
{
    QDir sourceDir = QDir(path);
    QString sourceName = sourceDir.dirName();
    if (sourceDir.exists() && sourceDir.entryInfoList().count()>2)
    {
        sourceDir.cdUp();
        for (int i = 0; ; i++)
        {
            QString newPath = sourceDir.absoluteFilePath(sourceName + "_" + QString::number(i, 10));
            if (!QDir(newPath).exists())
            {
                return newPath;
            }
        }
    }
    else
    {
        return path;
    }
}


QString PDFCreator::getFileName(QString path)
{
    int startIndex = getLastPathSeparatorIndex(path);
    if (path[path.count() - 1] == '\\' || path[path.count() - 1] == '/')
    {
        return path.mid(startIndex + 1, path.count() - startIndex - 2);
    }
    else
    {
        return path.mid(startIndex + 1, path.count() - startIndex - 1);
    }
}

// 保留尾部分隔符
QString PDFCreator::getFatherPath(QString path)
{
    int index = getLastPathSeparatorIndex(path);
    if (index < 0)
    {
        return QString();
    }

    return path.mid(0,index + 1);
}

QString PDFCreator::fixErrorPath(QString path)
{
    while (path.contains("//"))
    {
        path = path.replace("//", "/");
    }

    while (path.contains("\\\\"))
    {
        path = path.replace("\\\\", "\\");
    }

    return path;
}


int PDFCreator::getLastPathSeparatorIndex(QString path)
{
    int i = path.count() - 1;     //跳过可能存在的尾部'\\'或'/'
    while (true)
    {
        i--;
        if (i < 0)
        {
            // 出错，路径不合理
            return -1;
        }

        if (path[i] == '/' || path[i] == '\\')
        {
            break;
        }
    }

    return i;
}