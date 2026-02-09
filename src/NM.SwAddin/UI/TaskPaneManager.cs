using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using NM.Core;
using NM.Core.ProblemParts;
using NM.SwAddin.Pipeline;
using SolidWorks.Interop.sldworks;

namespace NM.SwAddin.UI
{
    /// <summary>
    /// Manages the lifecycle of the SolidWorks Task Pane for problem part review.
    /// Created in ConnectToSW, destroyed in DisconnectFromSW.
    /// The Task Pane tab is always present; content changes when problems are loaded.
    /// </summary>
    public sealed class TaskPaneManager : IDisposable
    {
        private ISldWorks _swApp;
        private ITaskpaneView _taskPaneView;
        private ProblemPartsTaskPaneControl _control;
        private bool _waitingForAction;
        private ProblemAction _lastAction;
        private Action<ProblemAction> _completionCallback;

        /// <summary>
        /// Whether the panel currently has problems loaded and is waiting for user action.
        /// Used by WorkflowDispatcher to poll for completion.
        /// </summary>
        public bool IsWaitingForAction => _waitingForAction;

        /// <summary>
        /// The action the user selected (Continue/Cancel). Valid after IsWaitingForAction becomes false.
        /// </summary>
        public ProblemAction LastAction => _lastAction;

        /// <summary>
        /// The list of problems that were fixed during the task pane session.
        /// </summary>
        public List<ProblemPartManager.ProblemItem> FixedProblems =>
            _control?.FixedProblems ?? new List<ProblemPartManager.ProblemItem>();

        /// <summary>
        /// Creates the Task Pane view in SolidWorks and hosts the ProblemPartsTaskPaneControl.
        /// Call once from ConnectToSW.
        /// </summary>
        public bool CreatePane(ISldWorks swApp)
        {
            _swApp = swApp;

            try
            {
                // CreateTaskpaneView2(iconPath, tooltip)
                // Empty string for icon = use default; SolidWorks will show a generic tab
                _taskPaneView = (ITaskpaneView)swApp.CreateTaskpaneView2(
                    string.Empty,
                    "NM Problem Parts");

                if (_taskPaneView == null)
                {
                    ErrorHandler.DebugLog("[TaskPane] CreateTaskpaneView2 returned null");
                    return false;
                }

                // AddControl hosts the COM-visible UserControl by its ProgId
                _control = (ProblemPartsTaskPaneControl)_taskPaneView.AddControl(
                    ProblemPartsTaskPaneControl.PROGID, "");

                if (_control == null)
                {
                    ErrorHandler.DebugLog("[TaskPane] AddControl returned null for ProgId: " + ProblemPartsTaskPaneControl.PROGID);
                    DestroyPane();
                    return false;
                }

                _control.Initialize(swApp);

                // Listen for user completing their review
                _control.ActionCompleted += OnActionCompleted;

                ErrorHandler.DebugLog("[TaskPane] Problem Parts task pane created successfully");
                return true;
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[TaskPane] Failed to create task pane: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Loads problems into the task pane and optionally shows the pane tab.
        /// </summary>
        /// <param name="problems">Problem items to display.</param>
        /// <param name="goodCount">Number of good models for display.</param>
        /// <param name="onComplete">Optional callback invoked when user completes review. If null, caller must poll IsWaitingForAction.</param>
        public void LoadProblems(List<ProblemPartManager.ProblemItem> problems, int goodCount, Action<ProblemAction> onComplete = null)
        {
            if (_control == null) return;

            _completionCallback = onComplete;
            _waitingForAction = true;
            _lastAction = ProblemAction.Cancel;
            _control.LoadProblems(problems, goodCount);

            ShowPane();
        }

        /// <summary>
        /// Makes the task pane tab visible/active.
        /// </summary>
        public void ShowPane()
        {
            try
            {
                _taskPaneView?.ShowView();
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[TaskPane] ShowView failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Clears problem data from the pane.
        /// </summary>
        public void ClearProblems()
        {
            _waitingForAction = false;
            _control?.Clear();
        }

        /// <summary>
        /// Destroys the Task Pane view. Call from DisconnectFromSW.
        /// </summary>
        public void DestroyPane()
        {
            try
            {
                if (_control != null)
                {
                    _control.ActionCompleted -= OnActionCompleted;
                    _control = null;
                }

                if (_taskPaneView != null)
                {
                    _taskPaneView.DeleteView();
                    Marshal.ReleaseComObject(_taskPaneView);
                    _taskPaneView = null;
                }

                ErrorHandler.DebugLog("[TaskPane] Problem Parts task pane destroyed");
            }
            catch (Exception ex)
            {
                ErrorHandler.DebugLog($"[TaskPane] DestroyPane error: {ex.Message}");
            }
        }

        /// <summary>
        /// Whether the task pane was successfully created.
        /// </summary>
        public bool IsCreated => _taskPaneView != null && _control != null;

        private void OnActionCompleted(object sender, ProblemAction action)
        {
            _lastAction = action;
            _waitingForAction = false;

            var callback = _completionCallback;
            _completionCallback = null;
            callback?.Invoke(action);
        }

        public void Dispose()
        {
            DestroyPane();
        }
    }
}
