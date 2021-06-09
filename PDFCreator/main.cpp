#include "PDFCreator.h"
#include <QtWidgets/QApplication>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    PDFCreator w;
    w.show();
    return a.exec();
}
