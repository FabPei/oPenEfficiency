using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using oPenEfficiency.Models;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Manages Undo and Redo operations using a Memento pattern with XML serialization.
    /// </summary>
    public class GanttUndoRedoManager
    {
        private readonly Stack<string> _undoStack = new Stack<string>();
        private readonly Stack<string> _redoStack = new Stack<string>();
        private string _currentStateXml;
        private bool _isPerformingUndoRedo = false;
        private int _transactionDepth;

        public bool CanUndo => _undoStack.Count > 0;
        public bool CanRedo => _redoStack.Count > 0;

        /// <summary>
        /// Saves a snapshot of the current state. Call this BEFORE making changes.
        /// </summary>
        public void SaveState(GanttProject project)
        {
            if (_isPerformingUndoRedo || _transactionDepth > 0 || project == null) return;

            string newXml = Serialize(project);

            // Avoid saving identical states
            if (newXml == _currentStateXml) return;

            if (!string.IsNullOrEmpty(_currentStateXml))
            {
                _undoStack.Push(_currentStateXml);
            }

            _currentStateXml = newXml;
            _redoStack.Clear(); // Clear redo on new action
        }

        /// <summary>
        /// Records a multi-step mutation as one undoable action. SaveState calls raised by
        /// collection/property events during the mutation are deliberately ignored.
        /// </summary>
        public void ExecuteTransaction(GanttProject project, Action mutation)
        {
            if (project == null || mutation == null) return;

            bool isOuterTransaction = _transactionDepth == 0;
            if (isOuterTransaction) SaveState(project);

            _transactionDepth++;
            try
            {
                mutation();
            }
            finally
            {
                _transactionDepth--;
                if (isOuterTransaction) SaveState(project);
            }
        }

        public GanttProject Undo(GanttProject currentProject)
        {
            if (!CanUndo) return currentProject;

            _isPerformingUndoRedo = true;
            try
            {
                _redoStack.Push(Serialize(currentProject));
                _currentStateXml = _undoStack.Pop();
                return Deserialize(_currentStateXml) ?? currentProject;
            }
            finally
            {
                _isPerformingUndoRedo = false;
            }
        }

        public GanttProject Redo(GanttProject currentProject)
        {
            if (!CanRedo) return currentProject;

            _isPerformingUndoRedo = true;
            try
            {
                _undoStack.Push(Serialize(currentProject));
                _currentStateXml = _redoStack.Pop();
                return Deserialize(_currentStateXml) ?? currentProject;
            }
            finally
            {
                _isPerformingUndoRedo = false;
            }
        }

        private string Serialize(GanttProject project)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(GanttProject));
                using (var stringWriter = new StringWriter())
                {
                    serializer.Serialize(stringWriter, project);
                    return stringWriter.ToString();
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttUndoRedoManager.Serialize");
                return null;
            }
        }

        private GanttProject Deserialize(string xml)
        {
            if (string.IsNullOrEmpty(xml)) return null;
            try
            {
                var serializer = new XmlSerializer(typeof(GanttProject));
                using (var stringReader = new StringReader(xml))
                {
                    var proj = (GanttProject)serializer.Deserialize(stringReader);
                    // Reattach collections
                    var tasks = proj.Tasks;
                    proj.Tasks = null;
                    proj.Tasks = tasks;
                    return proj;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "GanttUndoRedoManager.Deserialize");
                return null;
            }
        }
    }
}
