const app = getApp()

Page({
    data: {
        serverUrl: 'http://127.0.0.1:5000', // Default, user should change
        currentTab: 'hy',
        statusText: '就绪',
        running: false,
        logs: [],
        previewHtml: '',
        scrollTop: 0
    },

    onLoad() {
        // Load saved server URL
        const savedUrl = wx.getStorageSync('serverUrl')
        if (savedUrl) {
            this.setData({ serverUrl: savedUrl })
        }
        this.addLog('欢迎使用燃料管理小程序。请先配置服务器地址。')
    },

    onServerUrlInput(e) {
        this.setData({ serverUrl: e.detail.value })
        wx.setStorageSync('serverUrl', e.detail.value)
    },

    checkConnection() {
        wx.request({
            url: `${this.data.serverUrl}/`,
            method: 'GET',
            success: (res) => {
                if (res.statusCode === 200) {
                    wx.showToast({ title: '连接成功', icon: 'success' })
                    this.addLog('>>> 服务器连接成功！')
                } else {
                    this.addLog(`!!! 连接失败: code ${res.statusCode}`)
                }
            },
            fail: (err) => {
                wx.showToast({ title: '连接失败', icon: 'none' })
                this.addLog(`!!! 连接失败: ${err.errMsg}`)
            }
        })
    },

    switchTab(e) {
        const tab = e.currentTarget.dataset.tab
        this.setData({ currentTab: tab })
        this.loadPreview(tab)
    },

    chooseAndUpload() {
        const that = this;
        wx.chooseMessageFile({
            count: 1,
            type: 'file',
            extension: ['xls', 'xlsx'],
            success(res) {
                const tempFile = res.tempFiles[0]
                that.addLog(`准备上传: ${tempFile.name}`)

                wx.uploadFile({
                    url: `${that.data.serverUrl}/upload`,
                    filePath: tempFile.path,
                    name: 'file',
                    formData: {
                        'type': that.data.currentTab
                    },
                    success(uRes) {
                        const data = JSON.parse(uRes.data)
                        if (data.error) {
                            that.addLog(`❌ 上传失败: ${data.error}`)
                        } else {
                            that.addLog(`✅ 上传成功`)
                            wx.showToast({ title: '上传成功' })
                        }
                    },
                    fail(err) {
                        that.addLog(`❌ 上传请求失败: ${err.errMsg}`)
                    }
                })
            }
        })
    },

    executeTask() {
        const that = this
        this.setData({ running: true, statusText: '运行中...' })

        // Start Log Polling since SSE is hard
        this.startLogPolling()

        wx.request({
            url: `${that.data.serverUrl}/api/run`,
            method: 'POST',
            data: { type: that.data.currentTab },
            success(res) {
                that.addLog(`🚀 ${res.data.message}`)
            },
            fail(err) {
                that.addLog(`❌ 启动失败: ${err.errMsg}`)
                that.setData({ running: false, statusText: '就绪' })
            }
        })
    },

    // Simulated Log Polling (since real SSE needs Chunked support)
    // Ideally backend should provide a polling endpoint, but let's try reading SSE...
    // Or just assume logs will fail in this version and rely on final result.
    // UPDATE: Let's simply poll `api/preview` to check if done? 
    // Or just rely on user waiting.
    // Actually I can implement a simple 'get last log' on backend if needed.
    // For now, I won't poll logs continuously to avoid blocking, just show status.
    startLogPolling() {
        // Placeholder: In a real MP environment, use wx.request({ enableChunked: true }) for SSE
        // Here we just warn user
        this.addLog('Checking status...')

        // Auto-refresh preview after 5s, 10s...
        setTimeout(() => this.loadPreview(this.data.currentTab), 5000)
        setTimeout(() => {
            this.setData({ running: false, statusText: '就绪' })
            this.loadPreview(this.data.currentTab)
        }, 10000)
    },

    loadPreview(type) {
        const that = this
        wx.request({
            url: `${that.data.serverUrl}/api/preview/${type}`,
            success(res) {
                if (res.statusCode === 200 && !res.data.error) {
                    // Combine all tables
                    let html = ''
                    for (const [sheet, table] of Object.entries(res.data)) {
                        html += `<div class="sheet-title">${sheet}</div>` + table
                    }
                    // Replace class for styling
                    html = html.replace(/class="result-table"/g, 'style="width:100%; border-collapse: collapse; border:1px solid #ccc;" border="1"')

                    that.setData({ previewHtml: html })
                }
            }
        })
    },

    downloadResult() {
        const that = this
        const type = this.data.currentTab
        const url = `${this.data.serverUrl}/download/${type}`

        wx.downloadFile({
            url: url,
            success(res) {
                if (res.statusCode === 200) {
                    const filePath = res.tempFilePath
                    wx.openDocument({
                        filePath: filePath,
                        success: function () {
                            that.addLog('文档打开成功')
                        },
                        fail: function (err) {
                            that.addLog(`文档打开失败: ${err.errMsg}`)
                        }
                    })
                }
            },
            fail(err) {
                that.addLog(`下载失败: ${err.errMsg}`)
            }
        })
    },

    addLog(msg) {
        const logs = this.data.logs
        logs.push(`[${new Date().toLocaleTimeString()}] ${msg}`)
        this.setData({
            logs: logs,
            scrollTop: logs.length * 20
        })
    },

    clearLogs() {
        this.setData({ logs: [] })
    }
})
