#ifndef PDFCreator_H
#define PDFCreator_H

#define foreach

#include<QtWidgets/QWidget>
#include<QApplication>
#include<QLabel>
#include<QLineEdit>
#include<QpushButton>
#include<QGroupBox>
#include<QBoxLayout>
#include<QString>
#include<QFileDialog>
#include<QStandardPaths>
#include<QDir>
#include<QMessageBox>
#include<QDesktopServices>
#include<QCheckBox>
#include<QButtonGroup>
#include<QRadioButton>
#include<QTextEdit>
#include<QProgressBar>
#include<QColor>
#include<QStringList>
#include<QAxObject>
#include<QPrinter>
#include<QFuture>
#include<QtConcurrent>
#include<QtGui/QPainter>
#include<QMutex>
#include<QtPrintSupport/QPrinter>

//#include"PythonOffice.h"

class PDFCreator : public QWidget
{
    Q_OBJECT

public:
    PDFCreator(QWidget *parent = Q_NULLPTR);
    ~PDFCreator();

public: signals: void flushProgressBar(int value);
public: signals: void sendOutput(QString tip, QString contain, int color);

private:
    void InitGUI();
    void InitObjects();
    void connectSignalAndSlot();
    QString confirmFoldName(QString path);
    int getLastPathSeparatorIndex(QString path);
    QString getFileName(QString path);
    QString getFatherPath(QString path);    // 保留尾部分隔符
    QString fixErrorPath(QString path);
    void runProcess();                      // 实际处理文件
    void recursionDealFile(QString sourcePath, QString destinationPath);
    int countTotalWorks(QString sourcePath);
    QString getRealName(QString name);

    // 转化核心函数
    bool ppt_to_pdf(QString sourceFile, QString destinationFile);
    bool word_to_pdf(QString sourceFile, QString destinationFile);
    bool excel_to_pdf(QString sourceFile, QString destinationFile);
    bool picture_to_pdf(QString sourceFile, QString destinationFile);

private slots:
    void selectSourceFlod();                // 选择源文件夹
    void openSourceFold();                  // 打开源文件夹
    void selectDestinationFold();
    void autoComputeDestinationFold();
    void openDestinationFold();              // 打开源文件夹
    bool ensureFoldExitst(QString path);
    void outputLine(QString tip, QString contain, int color=Qt::black);
    void startDealFiles();
    void flushGUIProgressBar(int value);

private:
    QFuture<void> thread;
    bool isRun = false;
    bool isPause = false;
    QMutex lock;
    int totalFilesCount;
    int dealFilesCount;

private:
    QVBoxLayout* vBox = NULL;

    // 文件夹
    QGroupBox* foldPathGroupBox = NULL;
    QVBoxLayout* foldPathVBoxLayout = NULL;
    // 选择源文件夹
    QHBoxLayout* sourceHBoxLayout = NULL;
    QLabel* sourceTipLabel = NULL;
    QLineEdit* sourcePathLineEdit = NULL;
    QPushButton* sourceSelectFoldPushButton = NULL;
    QPushButton* sourceOpenPushButton = NULL;
    // 选择目标文件夹
    QHBoxLayout* destinationHBoxLayout = NULL;
    QLabel* destinationTipLabel = NULL;
    QLineEdit* destinationPathLineEdit = NULL;
    QPushButton* destinationSelectFoldPushButton = NULL;
    QPushButton* destinationPathAutoComputePushButton = NULL;
    QPushButton* destinationOpenPushButton = NULL;

    // 文件类型
    QGroupBox* fileTypeGroupBox = NULL;
    QHBoxLayout* fileTypeHBox = NULL;
    QCheckBox* pptCheckBox = NULL;
    QCheckBox* excelCheckBox = NULL;
    QCheckBox* wordCheckBox = NULL;
    QCheckBox* pictureCheckBox = NULL;

    // 输出设置
    QHBoxLayout* outputSettingHBoxLayout = NULL;
    // 子文件
    QGroupBox* subfoldersGroupBox = NULL;
    QVBoxLayout* subfoldersVBoxLayout = NULL;
    QButtonGroup* subfoldersButtonGroup = NULL;
    QRadioButton* enableSubfoldersRadioButton = NULL;
    QRadioButton* disableSubfoldersRadioButton = NULL;

    // 文件夹结构
    QGroupBox* foldStructureGroupBox = NULL;
    QVBoxLayout* foldStructureVBoxLayout = NULL;
    QButtonGroup* foldStructureButtonGroup = NULL;
    QRadioButton* enableFoldStructureRadioButton = NULL;
    QRadioButton* disableFoldStructureRadioButton = NULL;

    // 拷贝无关文件
    QGroupBox* ignoreGroupBox = NULL;
    QVBoxLayout* ignoreVBoxLayout = NULL;
    QButtonGroup* ignoreButtonGroup = NULL;
    QRadioButton* enableIgnoreRadioButton = NULL;
    QRadioButton* disableIgnoreRadioButton = NULL;

    // 程序输出
    QGroupBox* outputGroupBox = NULL;
    QHBoxLayout* outputHBoxLayout = NULL;
    QTextEdit* outputTextEdit = NULL;
    
    // 程序控制
    QHBoxLayout* runControlHBoxLayout = NULL;
    QProgressBar* runProgressBar = NULL;
    QLabel* runProgressTipLabel = NULL;
    QPushButton* runPushButton = NULL;

};

#endif