using System.Text;
using System.IO;
﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TrackLib;


namespace TrackManager.Words
{
    public delegate void ChangeProgress(int persents);

    public class WriteWordCoordinates
    {
        public  void WriteWordDoc(string fileName, TrackView trackView)
        {
            try
            {
                trackView.Lock();
                Track track = trackView.GetTrack();
                string templateName = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), "docTemplate.dotx");

                List<TrackPointExt> points = track.GetTrackPoints();
                List<List<TrackPointExt>> pointsInFiles = new List<List<TrackPointExt>>();

                for (int i = 0; i * pointsPerFile < points.Count; i++)
                {
                    int count = (points.Count - (i + 1) * pointsPerFile >= 0) ? pointsPerFile : points.Count - i * pointsPerFile;
                    pointsInFiles.Add(points.GetRange(i * pointsPerFile, count));
                }

                int currentPercent = -1;
                for (int i = 0; i < pointsInFiles.Count; i++)
                {
                    WordDocument document = new WordDocument(templateName, false);
                    List<TrackPointExt> tempPoints = pointsInFiles[i];
                    for (int j = 0; j < tempPoints.Count; j++)
                    {
                        // событие изменения процентов
                        int processPointsCount = 0;
                        for (int k = 0; k < i; i++)
                        {
                            processPointsCount += pointsInFiles[i].Count;
                        }

                            processPointsCount = (j + processPointsCount) * 100 / points.Count;
                            if (processPointsCount > currentPercent)
                            {
                                OnChanged(processPointsCount);
                                currentPercent = processPointsCount;

                            }
                        
                        //
                        string[] row = new string[6];
                        row[0] = j.ToString();
                        row[1] = tempPoints[j].Point.Y.ToString();
                        row[2] = tempPoints[j].Point.X.ToString();
                        row[3] = tempPoints[j].H.ToString();
                        row[4] = track.GetSpeed(j + i * pointsPerFile).ToString();
                        row[5] = tempPoints[j].Time.ToString();
                        document.AddDataToTable(1, row);
                    }
                    document.Save(fileName + i + ".doc");
                    document.CloseDocument();
                    document.Close();
                }
                trackView.Unlock();
            }
            catch (Exception ex)
            {
                trackView.Unlock();
                throw ex;
            }
        }

        private static void OnChanged(int persent)
        {
            changeProgress(persent);
        }
        
        /// <summary>
        /// Свойство возвращает и задает максимальное количесвто точек в файле
        /// </summary>
        public int PointsPerFile
        {
            get
            {
                return pointsPerFile;
            }
            set
            {
                pointsPerFile = value;
            }
        }

        public static event ChangeProgress changeProgress; //событие, которое показывает изменение процентов
        private int pointsPerFile = 100;
    }
}
