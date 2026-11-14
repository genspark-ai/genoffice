#pragma once

#include <QHash>
#include <QJsonArray>
#include <QJsonObject>
#include <QJsonValue>
#include <QSet>
#include <QWidget>
#include <functional>

class QLabel;
class QTimer;
class QWebEngineProfile;
class QWebEngineView;

namespace foraoffice {
class DocsQtPrinting;

struct DocsQtOptions {
    QString packageRoot;
    QString nodeExecutable;
    QString stateDirectory;
    QStringList readRoots;
    QStringList writeRoots;
    QString defaultSaveDirectory;
    QString initialPath;
    QJsonObject initialContent;
    /** Optional managed-document service configuration; callbacks use setServiceHandler. */
    QJsonObject documentHost;
    QStringList documentCheckers;
    QStringList attachments;
    QStringList knowledge;
    QString language = QStringLiteral("zh");
};

class DocsQtEditor;

class DocsQtChannel final : public QObject {
    Q_OBJECT
public:
    explicit DocsQtChannel(DocsQtEditor& editor);
    Q_INVOKABLE void request(const QString& message);
    Q_INVOKABLE void complete(const QString& message);
signals:
    void reply(QString message);
    void command(QString message);
    void event(QString name, QString arguments);
private:
    DocsQtEditor& editor_;
};

/** Native Docs container; document processing is in the SDK and owned Node service. */
class DocsQtEditor final : public QWidget {
    Q_OBJECT
public:
    using Completion = std::function<void(const QJsonObject&)>;
    using ActionHandler = std::function<void(const QJsonObject& action,
                                             const QJsonObject& grant, Completion done)>;
    /** Return a JSON-RPC string/null result, or an error; complete on the GUI thread. */
    using InstructionResolver = std::function<void(const QJsonObject& request, Completion done)>;
    using CreateWordHandler = std::function<void(const QJsonObject& content, Completion done)>;
    using OpenFileHandler = std::function<bool(const QString& path)>;
    /** Complete on the GUI thread with a JSON-RPC result or error member, including cancellation. */
    using ServiceHandler = std::function<void(const QJsonArray& arguments, Completion done)>;
    /** Call before constructing QApplication, as required by Qt WebEngine. */
    static void registerUrlSchemes();
    explicit DocsQtEditor(DocsQtOptions options, QWidget* parent = nullptr);
    ~DocsQtEditor() override;
    QWebEngineView* view() const { return view_; }
    void setActionHandler(ActionHandler handler) { actionHandler_ = std::move(handler); }
    /** An empty resolver preserves editor defaults. Replacing it cancels pending resolutions. */
    void setInstructionResolver(InstructionResolver resolver);
    /** Cancel one pending instruction by the request's stable action id. */
    void cancelInstructionResolution(const QString& actionId);
    void setCreateWordHandler(CreateWordHandler handler) { createWordHandler_ = std::move(handler); }
    void setOpenFileHandler(OpenFileHandler handler) { openFileHandler_ = std::move(handler); }
    /** An empty handler restores the default; a handler's cancellation/failure never falls back. */
    void setServiceHandler(const QString& service, ServiceHandler handler);
    void setServiceEnabled(const QString& service, bool enabled);
    /** Invoke a named public browser SDK operation; completion is a JSON-RPC result/error. */
    void callSdk(const QString& operation, const QJsonObject& arguments, Completion done);
    /** Host-controlled standard HTTP MCP for the current document; no AI action is required. */
    void startMcp(Completion done);
    void mcpState(Completion done);
    void stopMcp(Completion done);
    /** Trusted host grants, never exposed as renderer/Agent services. */
    void grantFile(const QString& path, bool read, bool write, Completion done);
    void revokeFile(const QString& path, Completion done);
    /** Replace authorized attachment/knowledge selections; paths require host read grants. */
    void configureResources(const QStringList& attachments, const QStringList& knowledge, Completion done);
    /** Trusted native menu entry; uses the editor's existing manual command handler. */
    void triggerMenuCommand(const QString& command, const QString& payload = {});
    QString nativeText(const QString& key) const;
signals:
    void ready();
    void failed(QString message);
    void documentCreated(QJsonObject outcome);
    /** Custom file-menu actions only. Built-in commands execute inside the editor. */
    void menuActionTriggered(QString actionId, QJsonObject context);
    /** Custom ribbon actions from tabs, appended groups and quick access. */
    void ribbonActionTriggered(QString actionId, QJsonObject context);
    /** Resolver work should stop; any later completion is ignored. */
    void instructionResolutionCanceled(QString actionId);
    /** Current SDK document state, or JSON null with no active document. Emitted after ready. */
    void documentStateChanged(QJsonValue state);
    /** Status and document identity only; obtain credentials through startMcp/mcpState. */
    void mcpStateChanged(QJsonObject state);
protected:
    void closeEvent(QCloseEvent* event) override;
private:
    Q_SIGNAL void backendSend(QJsonObject message);
    Q_SIGNAL void backendStopRequested();
    friend class DocsQtChannel;
    void fromRenderer(const QString& source);
    void fromRendererResult(const QString& source);
    void backendMessage(const QJsonObject& message);
    void sendBackend(const QJsonObject& message);
    void backendCall(const QString& method, const QJsonObject& params, Completion done);
    void rendererReply(const QJsonValue& id, const QJsonValue& result);
    void rendererError(const QJsonValue& id, const QString& message, const QJsonObject& errorData = {});
    void hostEvent(const QString& name, const QJsonArray& args = {});
    void serviceCall(const QJsonValue& id, const QString& method, const QJsonArray& args);
    void serviceNotify(const QString& method, const QJsonArray& args);
    void fail(const QString& message);
    void stopBackend();
    void acceptClose();
    void checkDocumentClose(std::function<void(bool)> done);
    void resolveCloseCheck(bool allowed);
    void requestCloseState();
    DocsQtOptions options_;
    QJsonObject nativeStrings_;
    QJsonObject mcpState_{{"status", "stopped"}};
    QHash<QString, Completion> pending_;
    QHash<QString, Completion> sdkPending_;
    QSet<QString> editorRequests_;
    QHash<QString, QJsonValue> instructionRequests_;
    DocsQtChannel* channel_ = nullptr;
    QWebEngineProfile* profile_ = nullptr;
    QWebEngineView* view_ = nullptr;
    QLabel* status_ = nullptr;
    ActionHandler actionHandler_;
    InstructionResolver instructionResolver_;
    CreateWordHandler createWordHandler_;
    OpenFileHandler openFileHandler_;
    QHash<QString, ServiceHandler> serviceHandlers_;
    QSet<QString> disabledServices_;
    DocsQtPrinting* printing_ = nullptr;
    bool creationReported_ = false;
    bool ready_ = false;
    bool backendReady_ = false;
    bool stopping_ = false;
    bool disposing_ = false;
    bool closeCheck_ = false;
    bool closeAllowed_ = false;
    std::function<void(bool)> closeDecision_;
    QTimer* closeTimer_ = nullptr;
    quint64 closeGeneration_ = 0;
};

} // namespace foraoffice
