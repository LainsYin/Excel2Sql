#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMap>
#include <QMainWindow>

class QProgressDialog;

namespace Ui {
class MainWindow;
}
class QSettings;
class QAxObject;
class QTextStream;
class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

    void initWidget();

private:
    void setStyleSheet();
    void getInitValue(const QString &title, QMap<QString, QString> &value);
    void errorMessageBox(const QString &error);
signals:


private slots:

    void on_comboBox_currentIndexChanged(const QString &arg1);
    
    void on_pushButton_preview_clicked();

    void on_pushButton_convertion_clicked();

    void on_deleteRow();

    void on_listWidget_customContextMenuRequested(const QPoint &pos);

private:
    void selectType(const QString &srcPath, const QString &destPath, const QString &work_sheet_name = NULL,  const bool &isCsv = false);


    void match(QTextStream &out);
    QStringList splitCSVLine(const QString &lineStr);
    void updatePath();
//    int getColumn(QAxObject *work_sheet);
    void insertSqlStr(QTextStream &out, QString &sqlStr);
    void insertFMSqlStr(QTextStream &out, QString &sqlStr);
    void updateSqlStr(QTextStream &out, QString &sqlStr);
    void deleteSqlStr(QTextStream &out, QString &sqlStr);
    void hotSqlStr(QTextStream &out, QString &sqlStr);
    bool initRowValue_excel(QAxObject *work_sheet);
    bool initRowValue_csv(const QString &path);
    QString getCellValue(QAxObject *work_sheet, int row, int column);

    bool isStringLength(const QString &str, const int &length); //没有超出字符 返回NULL， 否则返回错误信息
    bool isNumber(const QString &str);
    int analyzeRowData(const QStringList &row); /// -1 正确， 其他错误
    int analyzeActorRowData(const QStringList &row);

    ///true 正确 false 出错
    bool queryNumberAndLength(const QString &value, const int &length, const bool &isNull = false); ///true 不对 false 正确

    void insertLogSql();
    void setInfoText();
    bool isEmpty();


    void outputMessage(QtMsgType type, const QMessageLogContext &context, const QString &msg);

private:
    Ui::MainWindow *ui;

    QAxObject *_work_sheet;
    QString defaultPath;
    QString destFilePath;
    QSettings *initConfig;
    QMap<QString, QString> matchs, media, mp3, actor;
    QMap<QString, QString> songlist, medialist, hot;
    QMap<QString, QString> type;

    QMap<QString, int> info_count;

    QList<QStringList> rowList;
    QStringList fieldList;
    QStringList errorList;

    int m_add, m_update, m_dele;
    int a_add, a_update, a_dele;
    int p3_add, p3_update, p3_dele;
    int match_update, other_count;

    int row_start;
    int column_start;
    int row_count;
    int column_count;

    QMenu *menu;
    QAction *action;
};

#endif // MAINWINDOW_H
