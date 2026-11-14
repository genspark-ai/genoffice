#include "DocsQtEditor.hpp"
#include "DocsQtBackend.hpp"
#include "DocsQtPrinting.hpp"
#include <QBuffer>
#include <QApplication>
#include <QClipboard>
#include <QImageReader>
#include <QMimeData>
#include <QAbstractButton>
#include <QCloseEvent>
#include <QDesktopServices>
#include <QDir>
#include <QFile>
#include <QFileDialog>
#include <QFileInfo>
#include <QJsonDocument>
#include <QLabel>
#include <QLocale>
#include <QMessageBox>
#include <QPageSize>
#include <QMimeDatabase>
#include <QNetworkAccessManager>
#include <QNetworkReply>
#include <QNetworkRequest>
#include <QPointer>
#include <QTimer>
#include <QThread>
#include <QUuid>
#include <QUrlQuery>
#include <QVBoxLayout>
#include <QWebChannel>
#include <QWebEnginePage>
#include <QWebEnginePermission>
#include <QWebEngineProfile>
#include <QWebEngineScript>
#include <QWebEngineScriptCollection>
#include <QWebEngineSettings>
#include <QWebEngineUrlRequestJob>
#include <QWebEngineUrlScheme>
#include <QWebEngineUrlSchemeHandler>
#include <QWebEngineView>
#include <memory>
#include <utility>

namespace foraoffice {
namespace {
QString json(const QJsonObject& value) { return QString::fromUtf8(QJsonDocument(value).toJson(QJsonDocument::Compact)); }
QJsonObject errorResult(const QJsonValue& id, const QString& message, const QJsonObject& data = {}) {
    QJsonObject error{{"code", -32000}, {"message", message}};
    if (!data.isEmpty()) error.insert("data", data);
    return {{"jsonrpc", "2.0"}, {"id", id}, {"error", error}};
}
bool editorOrigin(const QUrl& url) {
    return url.scheme() == QStringLiteral("foraoffice-docs") && url.host() == QStringLiteral("editor");
}

class ResourceHandler final : public QWebEngineUrlSchemeHandler {
public:
    ResourceHandler(QString root, QObject* parent) : QWebEngineUrlSchemeHandler(parent), root_(std::move(root)), network_(this) {}
    void setAssets(const QJsonObject& assets) {
        assetUrl_ = QUrl(assets.value("url").toString());
        token_ = assets.value("token").toString().toUtf8();
    }
    void requestStarted(QWebEngineUrlRequestJob* job) override {
        const auto url = job->requestUrl();
        if (job->requestMethod() != "GET" || (!job->initiator().isEmpty() && !editorOrigin(job->initiator()))) {
            job->fail(QWebEngineUrlRequestJob::RequestDenied); return;
        }
        job->setAdditionalResponseHeaders({{"Access-Control-Allow-Origin", "foraoffice-docs://editor"}, {"Cache-Control", "no-store"}});
        if (url.scheme() == QStringLiteral("genoffice-docx-media")) {
            if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs asset: native request received");
            if (assetUrl_.isEmpty() || token_.isEmpty()) { job->fail(QWebEngineUrlRequestJob::RequestDenied); return; }
            auto target = assetUrl_;
            QUrlQuery query; query.addQueryItem(QStringLiteral("url"), url.toString(QUrl::FullyEncoded)); target.setQuery(query);
            QNetworkRequest request(target);
            request.setRawHeader("Authorization", "Bearer " + token_);
            request.setAttribute(QNetworkRequest::RedirectPolicyAttribute, QNetworkRequest::ManualRedirectPolicy);
            auto* response = network_.get(request);
            response->setParent(job);
            response->setReadBufferSize(1024 * 1024);
            const auto failed = std::make_shared<bool>(false);
            const auto payload = std::make_shared<QByteArray>();
            constexpr qint64 maxResourceBytes = 512LL * 1024 * 1024;
            QObject::connect(response, &QNetworkReply::metaDataChanged, job, [job, response, failed, payload] {
                if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs asset: headers status=%d available=%lld", response->attribute(QNetworkRequest::HttpStatusCodeAttribute).toInt(), static_cast<long long>(response->bytesAvailable()));
                if (*failed) return;
                const auto length = response->header(QNetworkRequest::ContentLengthHeader).toLongLong();
                if (response->attribute(QNetworkRequest::HttpStatusCodeAttribute).toInt() != 200 || length > maxResourceBytes) {
                    *failed = true; response->abort(); job->fail(QWebEngineUrlRequestJob::UrlNotFound); return;
                }
                if (length > 0) payload->reserve(length);
            });
            const auto readChunk = [job, response, failed, payload] {
                if (*failed) return;
                const auto chunk = response->readAll();
                if (payload->size() + chunk.size() > maxResourceBytes) {
                    *failed = true; response->abort(); job->fail(QWebEngineUrlRequestJob::RequestDenied); return;
                }
                payload->append(chunk);
            };
            QObject::connect(response, &QIODevice::readyRead, job, readChunk);
            QObject::connect(response, &QNetworkReply::finished, job, [job, response, failed, payload, readChunk] {
                if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs asset: network finished error=%d available=%lld atEnd=%d", int(response->error()), static_cast<long long>(response->bytesAvailable()), response->atEnd());
                readChunk();
                if (*failed) return;
                if (response->error() != QNetworkReply::NoError) { job->fail(QWebEngineUrlRequestJob::RequestFailed); return; }
                // WebEngine reads this device on its IO thread. Publish completed bytes;
                // the main-thread QNetworkReply must never be that reply device.
                auto* buffer = new QBuffer(job);
                buffer->setData(*payload);
                buffer->open(QIODevice::ReadOnly);
                job->reply(response->rawHeader("Content-Type"), buffer);
            });
            QObject::connect(job, &QObject::destroyed, response, [job, response] {
                // abort() can emit finished synchronously. The derived request job
                // is already destroyed here, before QObject disconnects its signals.
                QObject::disconnect(response, nullptr, job, nullptr);
                response->abort();
            });
            return;
        }
        if (!editorOrigin(url)) { job->fail(QWebEngineUrlRequestJob::RequestDenied); return; }
        const auto path = url.path();
        QString base;
        QString relative;
        if (path.startsWith(QStringLiteral("/sdk/"))) { base = root_ + "/embedded"; relative = path.mid(5); }
        else if (path == "/index.html" || path == "/bootstrap.js") { base = root_ + "/qt"; relative = path.mid(1); }
        else { job->fail(QWebEngineUrlRequestJob::UrlNotFound); return; }
        const auto canonicalBase = QFileInfo(base).canonicalFilePath();
        const auto filePath = QFileInfo(QDir(base).filePath(relative)).canonicalFilePath();
        const auto within = QDir(canonicalBase).relativeFilePath(filePath);
        if (canonicalBase.isEmpty() || filePath.isEmpty() || within == ".." || within.startsWith("../") || QDir::isAbsolutePath(within)) {
            job->fail(QWebEngineUrlRequestJob::RequestDenied); return;
        }
        auto* file = new QFile(filePath, job);
        if (!file->open(QIODevice::ReadOnly)) { job->fail(QWebEngineUrlRequestJob::UrlNotFound); return; }
        const auto mime = filePath.endsWith(".js") || filePath.endsWith(".mjs") ? QByteArray("text/javascript")
            : QMimeDatabase().mimeTypeForFile(filePath).name().toUtf8();
        job->reply(mime, file);
    }
private:
    QString root_;
    QNetworkAccessManager network_;
    QUrl assetUrl_;
    QByteArray token_;
};

class EditorPage final : public QWebEnginePage {
public:
    EditorPage(QWebEngineProfile* profile, QObject* parent) : QWebEnginePage(profile, parent) {}
protected:
    bool acceptNavigationRequest(const QUrl& url, NavigationType type, bool mainFrame) override {
        if (!mainFrame) return false;
        if (editorOrigin(url) && url.path() == "/index.html") return true;
        if (type == NavigationTypeLinkClicked && (url.scheme() == "https" || url.scheme() == "http" || url.scheme() == "mailto"))
            QDesktopServices::openUrl(url);
        return false;
    }
};
} // namespace

DocsQtChannel::DocsQtChannel(DocsQtEditor& editor) : QObject(&editor), editor_(editor) {}
void DocsQtChannel::request(const QString& message) { editor_.fromRenderer(message); }
void DocsQtChannel::complete(const QString& message) { editor_.fromRendererResult(message); }

void DocsQtEditor::registerUrlSchemes() {
    for (const auto& name : {QByteArray("foraoffice-docs"), QByteArray("genoffice-docx-media")}) {
        QWebEngineUrlScheme scheme(name);
        scheme.setSyntax(QWebEngineUrlScheme::Syntax::Host);
        scheme.setFlags(QWebEngineUrlScheme::SecureScheme | QWebEngineUrlScheme::CorsEnabled | QWebEngineUrlScheme::FetchApiAllowed);
        QWebEngineUrlScheme::registerScheme(scheme);
    }
}

DocsQtEditor::DocsQtEditor(DocsQtOptions options, QWidget* parent) : QWidget(parent), options_(std::move(options)) {
    closeTimer_ = new QTimer(this);
    closeTimer_->setSingleShot(true);
    connect(closeTimer_, &QTimer::timeout, this, [this] {
        resolveCloseCheck(false);
        status_->setText(nativeText("closeNoReplyMsg")); status_->show();
    });
    auto* layout = new QVBoxLayout(this);
    layout->setContentsMargins(0, 0, 0, 0);
    QFile localeFile(QDir(options_.packageRoot).filePath("qt/native-strings.json"));
    if (localeFile.open(QIODevice::ReadOnly)) {
        const auto locales = QJsonDocument::fromJson(localeFile.readAll()).object();
        if (!locales.contains(options_.language)) options_.language = QStringLiteral("en");
        nativeStrings_ = locales.value(options_.language).toObject();
    }
    setLayoutDirection(options_.language == "ar" || options_.language == "he" ? Qt::RightToLeft : Qt::LeftToRight);
    status_ = new QLabel(nativeText("dlgOpenDoc"), this);
    status_->setTextFormat(Qt::PlainText);
    status_->setWordWrap(true);
    layout->addWidget(status_);
    if (nativeStrings_.isEmpty()) {
        status_->setText(QStringLiteral("The packaged native translations are missing."));
        QTimer::singleShot(0, this, [this] { emit failed(status_->text()); });
        return;
    }
    profile_ = new QWebEngineProfile(this);
    auto* resources = new ResourceHandler(options_.packageRoot, profile_);
    profile_->installUrlSchemeHandler("foraoffice-docs", resources);
    profile_->installUrlSchemeHandler("genoffice-docx-media", resources);
    view_ = new QWebEngineView(this);
    auto* page = new EditorPage(profile_, view_);
    view_->setPage(page);
    printing_ = new DocsQtPrinting(view_, this);
    layout->addWidget(view_, 1);
    channel_ = new DocsQtChannel(*this);
    auto* webChannel = new QWebChannel(page);
    webChannel->registerObject(QStringLiteral("docsNative"), channel_);
    page->setWebChannel(webChannel);
    QFile library(QStringLiteral(":/qtwebchannel/qwebchannel.js"));
    if (!library.open(QIODevice::ReadOnly)) { fail(tr("The Qt WebChannel library is missing.")); return; }
    QWebEngineScript bootstrap;
    bootstrap.setName(QStringLiteral("foraoffice-docs-native-host"));
    bootstrap.setInjectionPoint(QWebEngineScript::DocumentCreation);
    bootstrap.setWorldId(QWebEngineScript::MainWorld);
    bootstrap.setRunsOnSubFrames(false);
    bootstrap.setSourceCode(QString::fromUtf8(library.readAll()) + QStringLiteral(
        "\nwindow.docsQtChannel=new Promise((resolve,reject)=>{"
        "if(!window.qt?.webChannelTransport){reject(new Error('The Qt document host is unavailable'));return;}"
        "new QWebChannel(window.qt.webChannelTransport,c=>resolve(c.objects.docsNative));});"));
    page->scripts().insert(bootstrap);
    page->settings()->setAttribute(QWebEngineSettings::JavascriptCanAccessClipboard, true);
    page->settings()->setAttribute(QWebEngineSettings::JavascriptCanPaste, true);
    connect(page, &QWebEnginePage::permissionRequested, this, [](QWebEnginePermission permission) {
        if (editorOrigin(permission.origin()) && (permission.permissionType() == QWebEnginePermission::PermissionType::ClipboardReadWrite ||
            permission.permissionType() == QWebEnginePermission::PermissionType::LocalFontsAccess)) permission.grant();
        else permission.deny();
    });
    connect(page, &QWebEnginePage::renderProcessTerminated, this, [this] { fail(tr("The Word renderer stopped. Your saved file is unchanged.")); });
    connect(page, &QWebEnginePage::windowCloseRequested, this, &QWidget::close);
    QJsonObject backendOptions{{"stateDirectory", options_.stateDirectory}, {"readRoots", QJsonArray::fromStringList(options_.readRoots)},
        {"writeRoots", QJsonArray::fromStringList(options_.writeRoots)}, {"defaultSaveDirectory", options_.defaultSaveDirectory},
        {"attachments", QJsonArray::fromStringList(options_.attachments)}, {"knowledge", QJsonArray::fromStringList(options_.knowledge)}};
    if (!options_.initialPath.isEmpty()) backendOptions.insert("initialPath", options_.initialPath);
    if (!options_.initialContent.isEmpty()) backendOptions.insert("initialContent", options_.initialContent);
    const auto cli = QDir(options_.packageRoot).filePath("cli/genoffice.cjs");
    if (QFileInfo::exists(cli)) backendOptions.insert("cli", QJsonObject{{"executable", options_.nodeExecutable}, {"entry", cli}});
    auto* thread = new QThread(qApp);
    auto* backend = new DocsQtBackend(options_.nodeExecutable,
        QDir(options_.packageRoot).filePath("services/node-host.cjs"), backendOptions);
    backend->moveToThread(thread);
    connect(thread, &QThread::started, backend, &DocsQtBackend::start);
    connect(this, &DocsQtEditor::backendSend, backend, &DocsQtBackend::send, Qt::QueuedConnection);
    connect(this, &DocsQtEditor::backendStopRequested, backend, &DocsQtBackend::stop, Qt::QueuedConnection);
    connect(backend, &DocsQtBackend::initialized, this, [this, resources](const QJsonObject& assets) {
        if (stopping_) return;
        backendReady_ = true;
        resources->setAssets(assets);
        view_->load(QUrl(QStringLiteral("foraoffice-docs://editor/index.html")));
    });
    connect(backend, &DocsQtBackend::received, this,
        [this](const QJsonObject& message, const DocsQtRequestLifetime& lifetime) {
            if (!stopping_ && (!lifetime || lifetime->load())) backendMessage(message);
        }, Qt::QueuedConnection);
    connect(backend, &DocsQtBackend::failed, this, [this](const QString& message) {
        if (!stopping_) fail(message);
    });
    connect(backend, &DocsQtBackend::stopped, thread, &QThread::quit, Qt::DirectConnection);
    connect(thread, &QThread::finished, backend, &QObject::deleteLater);
    connect(thread, &QThread::finished, thread, &QObject::deleteLater);
    // Normal view disposal never waits. Application exit joins cleanup, which needs no GUI replies.
    connect(qApp, &QCoreApplication::aboutToQuit, backend, &DocsQtBackend::stop, Qt::QueuedConnection);
    connect(qApp, &QCoreApplication::aboutToQuit, thread, [thread] { thread->wait(); });
    thread->start();
}

DocsQtEditor::~DocsQtEditor() {
    disposing_ = true;
    stopBackend();
    delete printing_;
    printing_ = nullptr;
    delete view_;
    view_ = nullptr;
    delete profile_;
    profile_ = nullptr;
}

void DocsQtEditor::sendBackend(const QJsonObject& message) {
    if (backendReady_ && !stopping_) emit backendSend(message);
}
void DocsQtEditor::backendCall(const QString& method, const QJsonObject& params, Completion done) {
    if (!backendReady_ || stopping_ || pending_.size() >= 64) {
        done(errorResult(QJsonValue::Null, tr("The document backend is unavailable or busy."))); return;
    }
    const auto id = QUuid::createUuid().toString(QUuid::WithoutBraces);
    pending_.insert(id, std::move(done));
    sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"method", method}, {"params", params}});
}
void DocsQtEditor::backendMessage(const QJsonObject& message) {
    if (message.value("jsonrpc") != "2.0") { fail(tr("Invalid document backend protocol.")); return; }
    const auto id = message.value("id").toString();
    const auto method = message.value("method").toString();
    if (method.isEmpty()) {
        auto callback = pending_.take(id);
        if (callback) callback(message);
    } else if (method == "host.createWord") {
        if (!createWordHandler_) {
            sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", QJsonObject{{"ok", false}, {"committed", false}, {"pending", true},
                {"error", "The native host has no Word view factory"}, {"errorCode", "unavailable"}}}});
            return;
        }
        createWordHandler_(message.value("params").toObject(), [self = QPointer<DocsQtEditor>(this), id](const QJsonObject& outcome) {
            if (self && !self->stopping_) self->sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", outcome}});
        });
    } else if (method == "host.renderHtml") {
        const auto html = message.value("params").toObject().value("html").toString();
        const QPageLayout layout(QPageSize(QPageSize::A4), QPageLayout::Portrait, QMarginsF(15.24, 15.24, 15.24, 15.24), QPageLayout::Millimeter);
        printing_->renderHtml(html, layout, [self = QPointer<DocsQtEditor>(this), id](DocsQtPrintResult result) {
            if (!self || self->stopping_) return;
            if (!result.ok) self->sendBackend(errorResult(id, result.error));
            else self->sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", QJsonObject{{"base64", QString::fromLatin1(result.pdf.toBase64())}}}});
        });
    } else if (method == "host.printDocument") {
        const auto params = message.value("params").toObject();
        const auto width = params.value("widthTwips");
        const auto height = params.value("heightTwips");
        const auto scale = params.value("scale");
        const auto layout = DocsQtPrinting::paperLayout(width.toDouble(), height.toDouble());
        if (!width.isDouble() || !height.isDouble() || !layout.isValid()
            || (!scale.isUndefined() && (!scale.isDouble() || scale.toDouble() != 1.0))) {
            sendBackend(errorResult(id, tr("Invalid document print size or scale.")));
            return;
        }
        if (!printing_) { sendBackend(errorResult(id, tr("The document print driver is unavailable."))); return; }
        printing_->printToPdf(layout, scale.isUndefined() ? 1.0 : scale.toDouble(),
            [self = QPointer<DocsQtEditor>(this), id](DocsQtPrintResult result) {
                if (!self || self->stopping_) return;
                if (!result.ok) self->sendBackend(errorResult(id, result.error));
                else self->sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", QJsonObject{{"base64", QString::fromLatin1(result.pdf.toBase64())}}}});
            });
    } else if (method == "host.openFile") {
        const auto path = message.value("params").toObject().value("path").toString();
        sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", openFileHandler_ ? openFileHandler_(path) : false}});
    } else if (method == "host.resourcesChanged") {
        hostEvent(QStringLiteral("onResourceStateChanged"), {message.value("params")});
    } else if (method == "host.mcpStateChanged") {
        mcpState_ = message.value("params").toObject();
        emit mcpStateChanged(mcpState_);
    } else if (method == "host.dialog") {
        const auto params = message.value("params").toObject();
        const auto kind = params.value("kind").toString();
        if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs native dialog: %s", qPrintable(kind));
        QJsonValue result;
        if (kind == "open" || kind == "image" || kind == "mergeSource") {
            const auto title = nativeText(kind == "image" ? "dlgInsertImage" : kind == "mergeSource" ? "dlgMergeDataSource" : "dlgOpenDoc");
            const auto filter = kind == "image" ? nativeText("filterImages") + " (*.png *.jpg *.jpeg *.gif)"
                : kind == "mergeSource" ? nativeText("filterMergeDataSource") + " (*.xlsx *.csv)" : nativeText("filterWord") + " (*.docx)";
            auto* picker = new QFileDialog(this, title, options_.defaultSaveDirectory, filter);
            picker->setFileMode(QFileDialog::ExistingFile);
            picker->setAttribute(Qt::WA_DeleteOnClose);
            connect(picker, &QDialog::finished, this, [this, id, picker](int outcome) {
                const auto paths = picker->selectedFiles();
                if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs native open: accepted=%d selected=%s", outcome == QDialog::Accepted, paths.isEmpty() ? "none" : qPrintable(QFileInfo(paths.front()).fileName()));
                const auto value = outcome == QDialog::Accepted && !paths.isEmpty() ? QJsonValue(paths.front()) : QJsonValue(QJsonValue::Null);
                sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", value}});
            });
            picker->QDialog::open();
            return;
        } else if (kind == "saveExport" || kind == "directory") {
            QString title = nativeText("dlgPickExportDir");
            QString filter;
            QString extension;
            QString defaultName;
            if (kind == "saveExport") {
                extension = params.value("extension").toString().toLower();
                defaultName = params.value("defaultName").toString();
                static const QSet<QString> extensions{"docx", "pdf", "html", "md", "png", "jpg", "jpeg", "gif", "webp", "bmp", "svg", "avif", "tif", "tiff"};
                if (!extensions.contains(extension) || defaultName.isEmpty()
                    || QFileInfo(defaultName).fileName() != defaultName || defaultName == "." || defaultName == ".."
                    || defaultName.contains(QChar(u'\0'))) {
                    sendBackend(errorResult(id, tr("Invalid export file name or extension.")));
                    return;
                }
                title = nativeText(extension == "pdf" ? "dlgExportPdf" : extension == "html" ? "dlgExportHtml"
                    : extension == "md" ? "dlgExportMarkdown" : "dlgSaveAs");
                const auto label = extension == "pdf" ? QStringLiteral("PDF") : extension == "html" ? QStringLiteral("HTML")
                    : extension == "md" ? QStringLiteral("Markdown") : extension == "docx" ? nativeText("filterWord") : nativeText("filterImages");
                filter = label + " (*." + extension + ")";
            }
            auto* picker = new QFileDialog(this, title, options_.defaultSaveDirectory, filter);
            picker->setAttribute(Qt::WA_DeleteOnClose);
            if (kind == "directory") {
                picker->setFileMode(QFileDialog::Directory);
                picker->setOption(QFileDialog::ShowDirsOnly);
            } else {
                picker->setAcceptMode(QFileDialog::AcceptSave);
                picker->setFileMode(QFileDialog::AnyFile);
                picker->setDefaultSuffix(extension);
                picker->selectFile(defaultName);
            }
            connect(picker, &QDialog::finished, this, [this, id, picker](int outcome) {
                const auto paths = picker->selectedFiles();
                const auto value = outcome == QDialog::Accepted && !paths.isEmpty() ? QJsonValue(paths.front()) : QJsonValue(QJsonValue::Null);
                sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", value}});
            });
            picker->QDialog::open();
            return;
        } else if (kind == "save") {
            const auto source = params.value("sourcePath").toString();
            const auto directory = source.isEmpty() ? options_.defaultSaveDirectory : QFileInfo(source).absolutePath();
            const auto path = QFileDialog::getSaveFileName(this, nativeText("dlgSaveAs"), QDir(directory).filePath(params.value("defaultName").toString()), nativeText("filterWord") + " (*.docx)");
            result = path.isEmpty() ? QJsonValue(QJsonValue::Null) : QJsonValue(path);
        } else if (kind == "recovery") {
            auto detail = nativeText("autosaveFoundBody");
            if (params.value("diskNewer").toBool()) detail += "\n\n" + nativeText("autosaveDiskNewer");
            QMessageBox question(QMessageBox::Question, nativeText("autosaveFoundTitle"), detail, QMessageBox::Yes | QMessageBox::No, this);
            question.button(QMessageBox::Yes)->setText(nativeText("autosaveRestore"));
            question.button(QMessageBox::No)->setText(nativeText("autosaveDiscard"));
            result = question.exec() == QMessageBox::Yes ? "restore" : "discard";
        } else if (kind == "overwrite") {
            QMessageBox question(QMessageBox::Question, nativeText("extModifiedMsg"), nativeText("extModifiedDetail"), QMessageBox::Yes | QMessageBox::No, this);
            question.button(QMessageBox::Yes)->setText(nativeText("btnOverwrite"));
            question.button(QMessageBox::No)->setText(nativeText("btnCancel"));
            result = question.exec() == QMessageBox::Yes;
        }
        else { sendBackend(errorResult(id, tr("Unknown native dialog."))); return; }
        sendBackend({{"jsonrpc", "2.0"}, {"id", id}, {"result", result}});
    } else if (method == "host.cancel") editorRequests_.remove(message.value("params").toObject().value("id").toString());
    else if (method.startsWith("editor.")) {
        if (!id.isEmpty()) editorRequests_.insert(id);
        emit channel_->command(json(message));
    }
}

void DocsQtEditor::rendererReply(const QJsonValue& id, const QJsonValue& result) {
    if (!id.isUndefined()) emit channel_->reply(json({{"jsonrpc", "2.0"}, {"id", id}, {"result", result}}));
}
void DocsQtEditor::rendererError(const QJsonValue& id, const QString& message, const QJsonObject& errorData) {
    if (!id.isUndefined()) emit channel_->reply(json(errorResult(id, message, errorData)));
}
void DocsQtEditor::hostEvent(const QString& name, const QJsonArray& args) {
    if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs host event: %s", qPrintable(name));
    emit channel_->event(name, QString::fromUtf8(QJsonDocument(args).toJson(QJsonDocument::Compact)));
}
void DocsQtEditor::triggerMenuCommand(const QString& command, const QString& payload) {
    if (ready_ && !stopping_) hostEvent(QStringLiteral("onMenuCommand"), {command, payload});
}
QString DocsQtEditor::nativeText(const QString& key) const { return nativeStrings_.value(key).toString(key); }
void DocsQtEditor::setServiceHandler(const QString& service, ServiceHandler handler) {
    Q_ASSERT(QThread::currentThread() == thread());
    if (handler) serviceHandlers_.insert(service, std::move(handler));
    else serviceHandlers_.remove(service);
}
void DocsQtEditor::setInstructionResolver(InstructionResolver resolver) {
    Q_ASSERT(QThread::currentThread() == thread());
    const auto requests = std::exchange(instructionRequests_, {});
    instructionResolver_ = std::move(resolver);
    QPointer<DocsQtEditor> self(this);
    for (auto request = requests.cbegin(); request != requests.cend(); ++request) {
        rendererError(request.value(), tr("The instruction resolution was canceled."), {{"errorCode", "canceled"}});
        if (!self) return;
        emit instructionResolutionCanceled(request.key());
        if (!self) return;
    }
}
void DocsQtEditor::cancelInstructionResolution(const QString& actionId) {
    Q_ASSERT(QThread::currentThread() == thread());
    const auto request = instructionRequests_.find(actionId);
    if (request == instructionRequests_.end()) return;
    const auto id = request.value();
    instructionRequests_.erase(request);
    QPointer<DocsQtEditor> self(this);
    rendererError(id, tr("The instruction resolution was canceled."), {{"errorCode", "canceled"}});
    if (!self) return;
    emit instructionResolutionCanceled(actionId);
}
void DocsQtEditor::setServiceEnabled(const QString& service, bool enabled) {
    Q_ASSERT(QThread::currentThread() == thread());
    if (enabled) disabledServices_.remove(service);
    else disabledServices_.insert(service);
}
void DocsQtEditor::callSdk(const QString& operation, const QJsonObject& arguments, Completion done) {
    Q_ASSERT(QThread::currentThread() == thread());
    if (!ready_ || stopping_ || sdkPending_.size() >= 64) {
        done(errorResult(QJsonValue::Null, tr("The document SDK is unavailable or busy."))); return;
    }
    const auto id = QUuid::createUuid().toString(QUuid::WithoutBraces);
    sdkPending_.insert(id, std::move(done));
    QTimer::singleShot(30000, this, [this, id] {
        auto callback = sdkPending_.take(id);
        if (callback) callback(errorResult(id, tr("The document SDK operation timed out.")));
    });
    emit channel_->command(json({{"jsonrpc", "2.0"}, {"id", id}, {"method", "sdk.request"},
        {"params", QJsonObject{{"operation", operation}, {"args", arguments}}}}));
}
void DocsQtEditor::grantFile(const QString& path, bool read, bool write, Completion done) {
    backendCall(QStringLiteral("files.grant"), {{"path", path}, {"read", read}, {"write", write}}, std::move(done));
}
void DocsQtEditor::configureResources(const QStringList& attachments, const QStringList& knowledge, Completion done) {
    backendCall(QStringLiteral("resources.configure"), {{"attachments", QJsonArray::fromStringList(attachments)},
        {"knowledge", QJsonArray::fromStringList(knowledge)}}, std::move(done));
}
void DocsQtEditor::startMcp(Completion done) {
    Q_ASSERT(QThread::currentThread() == thread());
    if (!ready_ || stopping_) {
        done(errorResult(QJsonValue::Null, tr("The live document editor is not ready."))); return;
    }
    backendCall(QStringLiteral("mcp.start"), {}, std::move(done));
}
void DocsQtEditor::mcpState(Completion done) {
    Q_ASSERT(QThread::currentThread() == thread());
    if (!backendReady_ || stopping_) { done({{"result", mcpState_}}); return; }
    backendCall(QStringLiteral("mcp.state"), {}, std::move(done));
}
void DocsQtEditor::stopMcp(Completion done) {
    Q_ASSERT(QThread::currentThread() == thread());
    if (!backendReady_ || stopping_) { done({{"result", mcpState_}}); return; }
    backendCall(QStringLiteral("mcp.stop"), {}, std::move(done));
}
void DocsQtEditor::revokeFile(const QString& path, Completion done) {
    backendCall(QStringLiteral("files.revoke"), {{"path", path}}, std::move(done));
}
void DocsQtEditor::fromRenderer(const QString& source) {
    const auto message = QJsonDocument::fromJson(source.toUtf8()).object();
    const auto id = message.value("id");
    if (message.value("jsonrpc") != "2.0") { rendererError(id, tr("Invalid native request.")); return; }
    const auto method = message.value("method").toString();
    const auto params = message.value("params").toObject();
    if (method == "services.call") serviceCall(id, params.value("method").toString(), params.value("args").toArray());
    else if (method == "services.notify") serviceNotify(params.value("method").toString(), params.value("args").toArray());
    else if (method == "editor.ready") { ready_ = true; status_->hide(); emit ready(); }
    else if (method == "editor.disposed") { ready_ = false; editorRequests_.clear(); stopBackend(); }
    else if (method == "sdk.ribbonAction") {
        if (ready_ && !stopping_) emit ribbonActionTriggered(params.value("actionId").toString(), params);
    }
    else if (method == "sdk.menuAction") {
        if (ready_ && !stopping_) emit menuActionTriggered(params.value("actionId").toString(), params);
    }
    else if (method == "sdk.documentStateChanged") {
        if (ready_ && !stopping_) emit documentStateChanged(params.value("state"));
    }
    else if (method == "editor.mcpRevoked") {
        sendBackend({{"jsonrpc", "2.0"}, {"method", "mcp.revoke"}, {"params", params}});
    }
    else if (method == "editor.cancelInstruction") {
        cancelInstructionResolution(params.value("id").toString());
    }
    else if (method == "editor.resolveInstruction") {
        if (!instructionResolver_) { rendererReply(id, QJsonValue::Null); return; }
        const auto resolver = instructionResolver_;
        const auto actionId = params.value("id").toString();
        if (actionId.isEmpty() || instructionRequests_.contains(actionId)) {
            rendererError(id, tr("The instruction resolution identity is invalid or already in use.")); return;
        }
        instructionRequests_.insert(actionId, id);
        resolver(params, [self = QPointer<DocsQtEditor>(this), id, actionId](const QJsonObject& response) {
            if (!self || self->stopping_ || self->instructionRequests_.value(actionId) != id) return;
            Q_ASSERT(QThread::currentThread() == self->thread());
            self->instructionRequests_.remove(actionId);
            if (response.contains("error")) {
                const auto error = response.value("error").toObject();
                self->rendererError(id, error.value("message").toString(), error.value("data").toObject());
            } else if (response.contains("result") && (response.value("result").isNull() || response.value("result").isString())) {
                self->rendererReply(id, response.value("result"));
            } else self->rendererError(id, self->tr("The instruction resolver returned an invalid result."));
        });
    }
    else if (method == "editor.submitAction") {
        if (!actionHandler_) { rendererError(id, tr("No external Harness conversation is configured.")); return; }
        backendCall(QStringLiteral("document.bindAction"), params, [this, id, params](const QJsonObject& response) {
            if (response.contains("error")) { rendererError(id, response.value("error").toObject().value("message").toString()); return; }
            actionHandler_(params, response.value("result").toObject(), [self = QPointer<DocsQtEditor>(this), id](const QJsonObject& result) {
                if (!self) return;
                if (result.value("ok").toBool()) self->rendererReply(id, QJsonValue::Null);
                else self->rendererError(id, result.value("error").toString("The external Harness rejected the action."));
            });
        });
    } else rendererError(id, tr("The native request is not available."));
}
void DocsQtEditor::fromRendererResult(const QString& source) {
    const auto message = QJsonDocument::fromJson(source.toUtf8()).object();
    const auto id = message.value("id").toString();
    if (message.value("jsonrpc") != "2.0") return;
    auto callback = sdkPending_.take(id);
    if (callback) { callback(message); return; }
    if (!editorRequests_.remove(id)) return;
    sendBackend(message);
}

void DocsQtEditor::serviceCall(const QJsonValue& id, const QString& method, const QJsonArray& args) {
    if (disabledServices_.contains(method)) {
        rendererError(id, tr("The host disabled this service."), {{"errorCode", "permission_denied"}});
        return;
    }
    const auto handler = serviceHandlers_.value(method);
    if (handler) {
        auto completed = std::make_shared<bool>(false);
        handler(args, [self = QPointer<DocsQtEditor>(this), id, completed](const QJsonObject& response) {
            if (!self || self->stopping_ || std::exchange(*completed, true)) return;
            Q_ASSERT(QThread::currentThread() == self->thread());
            if (response.contains("error")) {
                const auto error = response.value("error").toObject();
                self->rendererError(id, error.value("message").toString(), error.value("data").toObject());
            } else if (response.contains("result")) self->rendererReply(id, response.value("result"));
            else self->rendererError(id, tr("The host returned an invalid service outcome."));
        });
        return;
    }
    // A managed host cannot accidentally reach local persistence through an omitted callback.
    if (options_.documentHost.value("enabled").toBool()) {
        static const QSet<QString> localLifecycle{"openDocx", "openDocxPath", "openDocxDecrypt",
            "saveDocx", "saveDocxAs", "saveDocxNew", "saveDocxTo", "writeRecoveryCopy", "createDocument"};
        if (localLifecycle.contains(method)) {
            rendererError(id, tr("The managed host does not provide this document operation."), {{"errorCode", "unavailable"}});
            return;
        }
    }
    static const QSet<QString> nodeServices{"consumePendingOpenDocx", "consumeNewBlankDoc", "consumeAiDocContent", "consumeHeadlessExport",
        "openDocx", "openDocxPath", "openDocxDecrypt", "saveDocx", "saveDocxAs", "saveDocxNew", "saveDocxTo", "writeRecoveryCopy",
        "setDocPassword", "docPasswordIntentRevision", "discardDocPasswordIntents", "getRecentFiles", "fontMetrics", "fetchImage", "pickImage", "readImageSource",
        "createDocument", "getResourceState", "readAttachment", "kbList", "kbSearch", "kbRead", "runCli", "cancelCli", "pickMergeDataSource", "writeMergeOutput",
        "exportPdf", "exportHtml", "exportMarkdown", "printPdfBuffer", "saveMergedPdf", "pickExportImagesTarget", "takeExportPdf", "writeExportImage", "saveImageAs"};
    if (nodeServices.contains(method)) {
        backendCall(QStringLiteral("services.call"), {{"method", method}, {"args", args}}, [this, id, method](const QJsonObject& response) {
            if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs service completed: %s ok=%d", qPrintable(method), !response.contains("error"));
            if (response.contains("error")) rendererError(id, response.value("error").toObject().value("message").toString(), response.value("error").toObject().value("data").toObject());
            else if (method == "getResourceState") {
                auto result = response.value("result").toObject();
                if (!options_.documentHost.isEmpty()) result.insert("documentHost", options_.documentHost);
                result.insert("checkers", QJsonArray::fromStringList(options_.documentCheckers));
                rendererReply(id, result);
            } else rendererReply(id, response.value("result"));
        });
    } else if (method == "copyImageToClipboard") {
        const auto meta = args.at(1).toString();
        backendCall(QStringLiteral("services.call"), {{"method", "readClipboardImage"}, {"args", QJsonArray{args.at(0)}}},
            [this, id, meta](const QJsonObject& response) {
                const auto result = response.value("result").toObject();
                const auto encoded = result.value("base64").toString().toLatin1();
                const auto mime = result.value("mime").toString();
                if (response.contains("error") || !mime.startsWith("image/") || encoded.size() > 64 * 1024 * 1024) {
                    rendererReply(id, false); return;
                }
                const auto decoded = QByteArray::fromBase64Encoding(encoded, QByteArray::AbortOnBase64DecodingErrors);
                if (!decoded) { rendererReply(id, false); return; }
                QBuffer buffer;
                buffer.setData(decoded.decoded);
                buffer.open(QIODevice::ReadOnly);
                QImageReader reader(&buffer);
                const auto size = reader.size();
                if (!size.isValid() || qint64(size.width()) * size.height() > 64000000) { rendererReply(id, false); return; }
                const auto image = reader.read();
                if (image.isNull()) { rendererReply(id, false); return; }
                int width = image.width(), height = image.height();
                QString metaAttribute;
                if (!meta.isEmpty() && meta.size() <= 2048) {
                    QJsonParseError parseError;
                    const auto document = QJsonDocument::fromJson(meta.toUtf8(), &parseError);
                    if (parseError.error == QJsonParseError::NoError && document.isObject()) {
                        const auto attributes = document.object();
                        const auto displayWidth = attributes.value("imageWidthPx").toDouble();
                        const auto displayHeight = attributes.value("imageHeightPx").toDouble();
                        if (displayWidth > 0) width = qRound(qMin(displayWidth, 100000.0));
                        if (displayHeight > 0) height = qRound(qMin(displayHeight, 100000.0));
                        metaAttribute = " data-image-meta=\"" + meta.toHtmlEscaped() + "\"";
                    }
                }
                auto* payload = new QMimeData;
                payload->setImageData(image);
                payload->setHtml(QStringLiteral("<img src=\"data:%1;base64,%2\" width=\"%3\" height=\"%4\"%5>")
                    .arg(mime.toHtmlEscaped(), QString::fromLatin1(encoded)).arg(width).arg(height).arg(metaAttribute));
                QApplication::clipboard()->setMimeData(payload);
                rendererReply(id, true);
            });
    } else if (method == "listPrinters") {
        QJsonArray printers;
        for (const auto& printer : DocsQtPrinting::printers()) printers.append(QJsonObject{{"name", printer.name}, {"displayName", printer.displayName}});
        rendererReply(id, printers);
    } else if (method == "print") {
        const auto job = args.at(1).toObject();
        const auto paper = args.at(2).toObject();
        DocsQtPrintOptions options;
        if (job.value("deviceName").isString()) options.deviceName = job.value("deviceName").toString();
        options.copies = job.value("copies").toInt(1);
        options.collate = job.value("collate").toBool(true);
        const auto duplex = job.value("duplex").toString();
        options.duplex = duplex == "longEdge" ? DocsQtPrintOptions::Duplex::LongEdge : duplex == "shortEdge" ? DocsQtPrintOptions::Duplex::ShortEdge : DocsQtPrintOptions::Duplex::Simplex;
        options.layout = DocsQtPrinting::paperLayout(paper.value("widthTwips").toDouble(11906), paper.value("heightTwips").toDouble(16838));
        options.scale = args.isEmpty() ? 1.0 : args.at(0).toDouble(1.0);
        printing_->print(options, [self = QPointer<DocsQtEditor>(this), id](DocsQtPrintResult result) {
            if (self && !self->stopping_) self->rendererReply(id, QJsonObject{{"ok", result.ok}, {"error", result.error}, {"outcomeUnknown", result.outcomeUnknown}});
        });
    } else if (method == "getLanguage") rendererReply(id, options_.language);
    else if (method == "getSystemLocale") rendererReply(id, QLocale::system().bcp47Name());
    else if (method == "getTheme") rendererReply(id, "system");
    else if (method == "getDocumentTheme") rendererReply(id, "follow");
    else if (method == "getAutoSaveDefault") rendererReply(id, QJsonObject{{"on", false}, {"updatedAt", 0}});
    else if (method == "getAuthorName") rendererReply(id, qEnvironmentVariable("USERNAME", qEnvironmentVariable("USER")));
    else if (method == "getAiPanelPrefs") rendererReply(id, QJsonObject{{"side", "left"}, {"fontSize", "default"}, {"customFontSize", 14},
        {"spellcheck", true}, {"openInNewDocs", false}, {"showAiSidebar", false}});
    else if (method == "convertAltChunkHtml") rendererReply(id, QJsonValue::Null);
    else if (method == "confirmDocumentReplace") checkDocumentClose([this, id](bool allowed) { rendererReply(id, allowed); });
    else if (method == "spellLanguages" || method == "spellSetLanguages") {
        if (method == "spellSetLanguages") {
            QStringList languages; for (const auto value : args.at(0).toArray()) languages.append(value.toString());
            profile_->setSpellCheckLanguages(languages);
        }
        rendererReply(id, QJsonObject{{"active", QJsonArray::fromStringList(profile_->spellCheckLanguages())}, {"available", QJsonArray::fromStringList(profile_->spellCheckLanguages())}});
    } else if (method == "respellKick") { view_->setFocus(); rendererReply(id, QJsonValue::Null); }
    else if (method == "spellReplace") { view_->page()->replaceMisspelledWord(args.at(0).toString()); rendererReply(id, QJsonValue::Null); }
    else rendererError(id, tr("The Qt host does not provide this service: %1").arg(method));
}
void DocsQtEditor::serviceNotify(const QString& method, const QJsonArray& args) {
    if (method == "reportCreatedDocument") {
        const auto result = args.at(0).toObject();
        if (!creationReported_ && result.value("creationId") == options_.initialContent.value("creationId")) {
            creationReported_ = true;
            emit documentCreated(result.value("outcome").toObject());
        }
    } else if (method == "reportDocumentPath") {
        sendBackend({{"jsonrpc", "2.0"}, {"method", "document.hintPath"}, {"params", QJsonObject{{"path", args.at(0)}}}});
    } else if (method == "reportCloseCheck" && closeCheck_) {
        if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs close state: dirty=%d", args.at(0).toObject().value("dirty").toBool());
        closeTimer_->stop();
        const auto state = args.at(0).toObject();
        if (!state.value("dirty").toBool()) { resolveCloseCheck(true); return; }
        if (state.value("autoSave").toBool()) { hostEvent(QStringLiteral("onCloseSaveRequest")); return; }
        QMessageBox question(QMessageBox::Question, nativeText("closeUnsavedMsg"), nativeText("closeUnsavedDetail"),
            QMessageBox::Save | QMessageBox::Discard | QMessageBox::Cancel, this);
        question.button(QMessageBox::Save)->setText(nativeText("menuSave"));
        question.button(QMessageBox::Discard)->setText(nativeText("btnDontSave"));
        question.button(QMessageBox::Cancel)->setText(nativeText("btnCancel"));
        question.setDefaultButton(QMessageBox::Save);
        const auto choice = question.exec();
        if (choice == QMessageBox::Save) hostEvent(QStringLiteral("onCloseSaveRequest"));
        else if (choice == QMessageBox::Discard) resolveCloseCheck(true);
        else resolveCloseCheck(false);
    } else if (method == "reportCloseSaveResult" && closeCheck_) {
        if (args.at(0).toBool()) requestCloseState();
        else { resolveCloseCheck(false); status_->setText(nativeText("closeUnsavedMsg")); status_->show(); }
    }
}
void DocsQtEditor::fail(const QString& message) {
    mcpState_ = {{"status", "failed"}};
    status_->setText(nativeText("qtHostError").replace("{error}", message)); status_->show();
    const auto pending = std::move(pending_); pending_.clear();
    const auto sdkPending = std::move(sdkPending_); sdkPending_.clear();
    QPointer<DocsQtEditor> self(this);
    stopBackend();
    if (!self) return;
    for (const auto& callback : pending) {
        callback(errorResult(QJsonValue::Null, message));
        if (!self) return;
    }
    for (const auto& callback : sdkPending) {
        callback(errorResult(QJsonValue::Null, message));
        if (!self) return;
    }
    emit failed(message);
}
void DocsQtEditor::stopBackend() {
    if (stopping_) return;
    stopping_ = true;
    backendReady_ = false;
    ready_ = false;
    const auto pending = std::exchange(pending_, {});
    editorRequests_.clear();
    const auto sdkPending = std::exchange(sdkPending_, {});
    const auto instructionRequests = std::exchange(instructionRequests_, {});
    closeTimer_->stop();
    emit backendStopRequested();
    if (mcpState_.value("status") != "failed") mcpState_ = {{"status", "stopped"}};
    // QWidget::close normally hides the view; settle callers before their timers expire.
    // Destruction removes their owner and must not invoke reentrant application callbacks.
    if (disposing_) return;
    QPointer<DocsQtEditor> self(this);
    for (auto request = instructionRequests.cbegin(); request != instructionRequests.cend(); ++request) {
        rendererError(request.value(), tr("The instruction resolution was canceled."), {{"errorCode", "canceled"}});
        if (!self) return;
        emit instructionResolutionCanceled(request.key());
        if (!self) return;
    }
    for (const auto& callback : pending) {
        callback(errorResult(QJsonValue::Null, tr("The document backend operation was canceled.")));
        if (!self) return;
    }
    for (const auto& callback : sdkPending) {
        callback(errorResult(QJsonValue::Null, tr("The document SDK operation was canceled.")));
        if (!self) return;
    }
    emit mcpStateChanged(mcpState_);
}
void DocsQtEditor::acceptClose() { closeAllowed_ = true; close(); }
void DocsQtEditor::resolveCloseCheck(bool allowed) {
    closeTimer_->stop();
    closeCheck_ = false;
    auto done = std::move(closeDecision_); closeDecision_ = {};
    if (done) done(allowed);
}
void DocsQtEditor::checkDocumentClose(std::function<void(bool)> done) {
    if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs close check: ready=%d pending=%d visible=%d", ready_, closeCheck_, isVisible());
    if (!ready_) { done(true); return; }
    if (closeCheck_) { done(false); return; }
    closeCheck_ = true;
    const auto generation = ++closeGeneration_;
    closeDecision_ = std::move(done);
    backendCall(QStringLiteral("document.releaseAllActions"), {}, [this, generation](const QJsonObject& response) {
        if (qEnvironmentVariableIsSet("FORAOFFICE_DOCS_DIAGNOSTICS")) qInfo("Docs revoke completed: error=%d", response.contains("error"));
        if (!closeCheck_ || generation != closeGeneration_) return;
        if (response.contains("error")) { resolveCloseCheck(false); return; }
        requestCloseState();
    });
}
void DocsQtEditor::requestCloseState() {
    closeTimer_->start(30000);
    hostEvent(QStringLiteral("onCloseCheck"));
}
void DocsQtEditor::closeEvent(QCloseEvent* event) {
    if (closeAllowed_ || !ready_) { stopBackend(); event->accept(); return; }
    event->ignore();
    if (!closeCheck_) checkDocumentClose([this](bool allowed) { if (allowed) acceptClose(); });
}

} // namespace foraoffice
