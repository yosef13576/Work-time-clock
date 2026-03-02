import { useState, useEffect, useRef } from 'react'
import * as XLSX from 'xlsx'

function App() {
  const [clients, setClients] = useState(() => {
    try { return JSON.parse(localStorage.getItem('wtc_clients')) || [] } catch { return [] }
  })
  const [newClientName, setNewClientName] = useState('')
  const [selectedClientId, setSelectedClientId] = useState('')
  const [sessions, setSessions] = useState(() => {
    try { return JSON.parse(localStorage.getItem('wtc_sessions')) || [] } catch { return [] }
  })
  const [currentNote, setCurrentNote] = useState('')
  const [isRunning, setIsRunning] = useState(false)
  const [isPaused, setIsPaused] = useState(false)
  const [elapsedMs, setElapsedMs] = useState(0)
  const [sessionStartTime, setSessionStartTime] = useState(null)
  const [filterClient, setFilterClient] = useState('all')

  const intervalRef = useRef(null)
  const elapsedBeforePauseRef = useRef(0)
  const timerStartRef = useRef(null)

  useEffect(() => {
    localStorage.setItem('wtc_clients', JSON.stringify(clients))
  }, [clients])

  useEffect(() => {
    localStorage.setItem('wtc_sessions', JSON.stringify(sessions))
  }, [sessions])

  useEffect(() => {
    if (isRunning && !isPaused) {
      intervalRef.current = setInterval(() => {
        setElapsedMs(elapsedBeforePauseRef.current + (Date.now() - timerStartRef.current))
      }, 100)
    } else {
      clearInterval(intervalRef.current)
    }
    return () => clearInterval(intervalRef.current)
  }, [isRunning, isPaused])

  const addClient = () => {
    const name = newClientName.trim()
    if (!name) return
    if (clients.some(c => c.name.toLowerCase() === name.toLowerCase())) {
      alert('לקוח בשם זה כבר קיים')
      return
    }
    setClients(prev => [...prev, { id: Date.now().toString(), name }])
    setNewClientName('')
  }

  const removeClient = (id) => {
    setClients(prev => prev.filter(c => c.id !== id))
    if (selectedClientId === id) setSelectedClientId('')
  }

  const startTimer = () => {
    if (!selectedClientId) return
    setSessionStartTime(new Date())
    timerStartRef.current = Date.now()
    elapsedBeforePauseRef.current = 0
    setElapsedMs(0)
    setIsRunning(true)
    setIsPaused(false)
  }

  const togglePause = () => {
    if (isPaused) {
      timerStartRef.current = Date.now()
      setIsPaused(false)
    } else {
      elapsedBeforePauseRef.current = elapsedMs
      setIsPaused(true)
    }
  }

  const cancelTimer = () => {
    if (!confirm('לבטל את הסשן הנוכחי? הזמן לא יישמר.')) return
    clearInterval(intervalRef.current)
    setIsRunning(false)
    setIsPaused(false)
    setElapsedMs(0)
    elapsedBeforePauseRef.current = 0
    setCurrentNote('')
    setSessionStartTime(null)
  }

  const stopTimer = () => {
    clearInterval(intervalRef.current)
    const client = clients.find(c => c.id === selectedClientId)
    if (!client) return

    const endTime = new Date()
    const session = {
      id: Date.now().toString(),
      clientId: selectedClientId,
      clientName: client.name,
      date: sessionStartTime.toLocaleDateString('he-IL'),
      startTime: sessionStartTime.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' }),
      endTime: endTime.toLocaleTimeString('he-IL', { hour: '2-digit', minute: '2-digit' }),
      durationMs: elapsedMs,
      durationFormatted: formatTime(elapsedMs),
      notes: currentNote,
      dateISO: sessionStartTime.toISOString(),
    }

    setSessions(prev => [session, ...prev])
    setIsRunning(false)
    setIsPaused(false)
    setElapsedMs(0)
    elapsedBeforePauseRef.current = 0
    setCurrentNote('')
    setSessionStartTime(null)
  }

  const deleteSession = (id) => {
    if (!confirm('למחוק סשן זה?')) return
    setSessions(prev => prev.filter(s => s.id !== id))
  }

  const formatTime = (ms) => {
    const total = Math.floor(ms / 1000)
    const h = Math.floor(total / 3600)
    const m = Math.floor((total % 3600) / 60)
    const s = total % 60
    return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`
  }

  const msToHours = (ms) => (ms / 3600000).toFixed(2)

  const getTotalsPerClient = () => {
    const totals = {}
    sessions.forEach(s => {
      totals[s.clientName] = (totals[s.clientName] || 0) + s.durationMs
    })
    return totals
  }

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new()

    const filteredSessions = filterClient === 'all'
      ? sessions
      : sessions.filter(s => s.clientId === filterClient)

    const sessionsData = filteredSessions.map(s => ({
      'לקוח': s.clientName,
      'תאריך': s.date,
      'שעת התחלה': s.startTime,
      'שעת סיום': s.endTime,
      'משך (שעות)': parseFloat(msToHours(s.durationMs)),
      'משך מפורמט': s.durationFormatted,
      'הערות': s.notes || '',
    }))

    const ws = XLSX.utils.json_to_sheet(sessionsData)
    ws['!cols'] = [
      { wch: 20 }, { wch: 14 }, { wch: 12 }, { wch: 12 },
      { wch: 14 }, { wch: 12 }, { wch: 30 },
    ]
    XLSX.utils.book_append_sheet(wb, ws, 'סשני עבודה')

    const totals = getTotalsPerClient()
    const summaryData = Object.entries(totals).map(([name, ms]) => ({
      'לקוח': name,
      'סה"כ שעות': parseFloat(msToHours(ms)),
      'מספר סשנים': sessions.filter(s => s.clientName === name).length,
    }))
    const wsSummary = XLSX.utils.json_to_sheet(summaryData)
    wsSummary['!cols'] = [{ wch: 20 }, { wch: 14 }, { wch: 14 }]
    XLSX.utils.book_append_sheet(wb, wsSummary, 'סיכום לפי לקוח')

    const dateStr = new Date().toLocaleDateString('he-IL').replace(/\//g, '-')
    XLSX.writeFile(wb, `דוח-שעות-${dateStr}.xlsx`)
  }

  const clientTotals = getTotalsPerClient()
  const selectedClient = clients.find(c => c.id === selectedClientId)
  const totalAllMs = sessions.reduce((sum, s) => sum + s.durationMs, 0)

  const filteredSessions = filterClient === 'all'
    ? sessions
    : sessions.filter(s => s.clientId === filterClient)

  return (
    <div className="app" dir="rtl">
      <header className="app-header">
        <div className="header-content">
          <h1>⏱ מעקב שעות עבודה</h1>
          <p>ניהול שעות עבודה לפרילנסרים</p>
        </div>
        {sessions.length > 0 && (
          <div className="header-total">
            <span className="total-label">סה&quot;כ שעות</span>
            <span className="total-value">{msToHours(totalAllMs)}</span>
          </div>
        )}
      </header>

      <main>
        {/* Client Management */}
        <section className="card">
          <h2>לקוחות</h2>
          <div className="add-client-row">
            <input
              type="text"
              className="input"
              placeholder="שם הלקוח החדש..."
              value={newClientName}
              onChange={e => setNewClientName(e.target.value)}
              onKeyDown={e => e.key === 'Enter' && addClient()}
              disabled={isRunning}
            />
            <button onClick={addClient} className="btn btn-primary" disabled={isRunning}>
              + הוסף לקוח
            </button>
          </div>
          <div className="chips-row">
            {clients.map(c => (
              <div key={c.id} className="chip">
                <span>{c.name}</span>
                {!isRunning && (
                  <button
                    className="chip-remove"
                    onClick={() => removeClient(c.id)}
                    title="הסר לקוח"
                  >×</button>
                )}
                {clientTotals[c.name] && (
                  <span className="chip-hours">{msToHours(clientTotals[c.name])}ש׳</span>
                )}
              </div>
            ))}
            {clients.length === 0 && (
              <p className="empty-hint">הוסף לקוחות כדי להתחיל לעקוב אחרי שעות</p>
            )}
          </div>
        </section>

        {/* Timer */}
        <section className={`card timer-card ${isRunning ? (isPaused ? 'state-paused' : 'state-running') : ''}`}>
          <h2>טיימר עבודה</h2>

          {!isRunning ? (
            <div className="select-row">
              <label className="select-label">בחר לקוח לעבודה:</label>
              <select
                className="select"
                value={selectedClientId}
                onChange={e => setSelectedClientId(e.target.value)}
              >
                <option value="">-- בחר לקוח --</option>
                {clients.map(c => (
                  <option key={c.id} value={c.id}>{c.name}</option>
                ))}
              </select>
            </div>
          ) : (
            <div className="working-for">
              עובד עבור: <strong>{selectedClient?.name}</strong>
              {isPaused && <span className="pause-badge">⏸ הפסקה</span>}
            </div>
          )}

          <div className={`timer-display ${isPaused ? 'paused' : ''}`}>
            {formatTime(elapsedMs)}
          </div>

          <div className="timer-buttons">
            {!isRunning ? (
              <button
                className="btn btn-start"
                onClick={startTimer}
                disabled={!selectedClientId}
              >
                ▶ התחל עבודה
              </button>
            ) : (
              <>
                <button
                  className={`btn ${isPaused ? 'btn-resume' : 'btn-pause'}`}
                  onClick={togglePause}
                >
                  {isPaused ? '▶ המשך עבודה' : '⏸ הפסקה'}
                </button>
                <button className="btn btn-stop" onClick={stopTimer}>
                  ⏹ סיים ושמור
                </button>
                <button className="btn btn-cancel" onClick={cancelTimer}>
                  ✕ בטל סשן
                </button>
              </>
            )}
          </div>

          {isRunning && (
            <textarea
              className="notes-input"
              placeholder="הערות לסשן זה (אופציונלי) — תיאור המשימה, פרוייקט..."
              value={currentNote}
              onChange={e => setCurrentNote(e.target.value)}
            />
          )}
        </section>

        {/* Summary per client */}
        {Object.keys(clientTotals).length > 0 && (
          <section className="card">
            <h2>סיכום שעות לפי לקוח</h2>
            <div className="summary-grid">
              {Object.entries(clientTotals)
                .sort((a, b) => b[1] - a[1])
                .map(([name, ms]) => (
                  <div key={name} className="summary-box">
                    <div className="summary-client">{name}</div>
                    <div className="summary-hours">{msToHours(ms)}</div>
                    <div className="summary-label">שעות</div>
                    <div className="summary-sessions">
                      {sessions.filter(s => s.clientName === name).length} סשנים
                    </div>
                  </div>
                ))}
            </div>
          </section>
        )}

        {/* Sessions Log */}
        <section className="card">
          <div className="log-header">
            <h2>יומן עבודה</h2>
            <div className="log-actions">
              {clients.length > 0 && sessions.length > 0 && (
                <select
                  className="select select-sm"
                  value={filterClient}
                  onChange={e => setFilterClient(e.target.value)}
                >
                  <option value="all">כל הלקוחות</option>
                  {clients.map(c => (
                    <option key={c.id} value={c.id}>{c.name}</option>
                  ))}
                </select>
              )}
              {sessions.length > 0 && (
                <button className="btn btn-export" onClick={exportToExcel}>
                  📊 ייצוא לאקסל
                </button>
              )}
            </div>
          </div>

          {filteredSessions.length === 0 ? (
            <p className="empty-hint" style={{ padding: '2rem 0' }}>
              {sessions.length === 0
                ? 'אין סשנים מוקלטים עדיין. התחל לעבוד!'
                : 'אין סשנים עבור הלקוח שנבחר'}
            </p>
          ) : (
            <div className="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>לקוח</th>
                    <th>תאריך</th>
                    <th>התחלה</th>
                    <th>סיום</th>
                    <th>משך</th>
                    <th>שעות</th>
                    <th>הערות</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {filteredSessions.map(s => (
                    <tr key={s.id}>
                      <td><span className="client-tag">{s.clientName}</span></td>
                      <td>{s.date}</td>
                      <td>{s.startTime}</td>
                      <td>{s.endTime}</td>
                      <td className="duration-cell"><strong>{s.durationFormatted}</strong></td>
                      <td className="hours-cell">{msToHours(s.durationMs)}</td>
                      <td className="notes-cell" title={s.notes}>{s.notes || '—'}</td>
                      <td>
                        <button
                          className="btn-icon"
                          onClick={() => deleteSession(s.id)}
                          title="מחק סשן"
                        >🗑</button>
                      </td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  <tr>
                    <td colSpan="5" className="total-row-label">סה&quot;כ ({filteredSessions.length} סשנים)</td>
                    <td className="total-row-value">
                      {msToHours(filteredSessions.reduce((s, r) => s + r.durationMs, 0))} ש׳
                    </td>
                    <td colSpan="2"></td>
                  </tr>
                </tfoot>
              </table>
            </div>
          )}
        </section>
      </main>

      <footer className="app-footer">
        <p>כל הנתונים נשמרים באופן מקומי בדפדפן שלך</p>
      </footer>
    </div>
  )
}

export default App
