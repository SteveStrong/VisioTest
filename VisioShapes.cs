using System;
using System.Collections.Generic;

namespace VisioShapeExtractor
{
    /// <summary>
    /// Container for all shapes extracted from a Visio document
    /// </summary>
    public class VisioShapes
    {
        /// <summary>
        /// The filename or path of the source Visio document
        /// </summary>
        public string filename { get; set; } = string.Empty;
        
        /// <summary>
        /// List of 2D shapes (normal shapes, rectangles, circles, etc.)
        /// </summary>
        public List<Shape2D> Shape2D { get; set; } = new();
        
        /// <summary>
        /// List of 1D shapes (connectors, lines, etc.)
        /// </summary>
        public List<Shape1D> Shape1D { get; set; } = new();
    }
}
