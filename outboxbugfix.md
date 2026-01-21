Outbox delay rule bug fix notes

Context
- Symptom: inline replies get stuck in Outbox with NONE date, non-italic, never send.
- Trigger: 2-minute delay rule on send.
- Root cause: any COM access to the inline draft during the delay-rule window creates a lock that prevents Outlook from sending.

What actually fixed it
- Do not touch the inline draft COM object at all during inline compose.
- Do not read ActiveInlineResponse in SelectionChange (it returns the draft COM object).
- Use Application.ItemSend to start a full 2-minute cooldown to block all selection handling while the delay rule is active.

Exact changes applied
1) Explorer_InlineResponse
   - Removed all reads from the inline draft (Subject, Submitted, EntryID, etc.).
   - Only flips UI into compose mode through a new method that does not take a MailItem.
2) Explorer_SelectionChange
   - If compose mode is active, return immediately.
   - No ActiveInlineResponse access.
3) Application.ItemSend
   - Deactivate compose mode.
   - Start a 2-minute cooldown (_ignoreSelectionUntil) to prevent selection logic from touching anything while the delay rule runs.

Why the earlier fixes were insufficient
- Guarding with ActiveInlineResponse still reads the inline draft COM object.
- Releasing COM objects on some early-return paths is good hygiene, but it does not prevent the lock if the draft is touched at all.

What is currently broken by the fix
1) Inline compose cancel detection
   - Without ActiveInlineResponse checks, the UI remains in compose mode if the user cancels an inline reply.
   - Current behavior: compose mode exits only on send (ItemSend).
2) Inline deferred scheduling
   - Inline compose no longer loads or saves deferred properties because we do not touch the draft.
   - Popup Inspector compose still works as before.

How to restore inline features safely
Goal: re-enable inline features without touching the draft during the delay-rule window.

Option A (safe, minimal risk)
- Add a manual "Exit compose mode" action in the pane.
- Keep inline compose "read-only" (no draft properties access).
- Keep all draft interaction limited to popup Inspector compose.

Option B (automated exit with controlled risk)
- Add a timer that checks ActiveInlineResponse to detect cancel.
- Risk: any access to ActiveInlineResponse returns the draft COM object and can reintroduce the lock.
- If you do this, restrict it to periods when the delay rule is disabled.

Option C (hybrid, likely workable)
- Allow inline features only after send completes.
- Store any deferred actions in memory during inline compose and apply them in SentItems_ItemAdd.
- This avoids draft access while the item is queued.

Keepers to retain (recommended)
1) PropertyAccessor for UserProperties (stealth mode)
   - Rationale: UserProperties access is heavier and can produce locks/security prompts.
   - Keep using PropertyAccessor for reading deferred data.
2) Outbox guard
   - Always skip items whose parent folder is Outbox.
   - Avoids touching queued items, which is a high-risk area.
3) Explicit Marshal.ReleaseComObject in hot paths
   - Important for inline events and frequent Explorer events.
   - Reduces chances of lingering COM locks.
4) Cooldown timer around send
   - This prevents post-send selection logic from touching conversation parent items during the delay rule.

Testing checklist
- Inline reply with 2-minute delay rule enabled.
- Send and confirm Outbox shows italic item with scheduled time.
- Item leaves Outbox after 2 minutes.
- Verify selection handling resumes after cooldown.

Files to revisit if re-enabling inline features
- FastPMHelperAddin/ThisAddIn.cs (Explorer_InlineResponse, Explorer_SelectionChange, Application_ItemSend)
- FastPMHelperAddin/UI/ProjectActionPane.xaml.cs (OnInlineComposeActivated, compose-mode UI)
