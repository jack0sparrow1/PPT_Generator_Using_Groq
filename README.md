# AI-Powered Presentation Generator - Learning Journey

## What I Think I Learned Building This Project

### 1. API Integration & API Keys
- Integrated with Groq's LLM API using their Python SDK for content generation
- Implemented image search functionality using Pexels API with HTTP requests
- Used environment variables for API key management (learned hardcoding is not secure)
- Gained awareness of API rate limits and implemented basic error handling

### 2. Presentation Automation with python-pptx
- Programmatically created PowerPoint slides using the pptx library
- Managed slide layouts, dimensions, margins, and element positioning using Inches
- Applied text formatting: font sizes (Pt), bold styling, alignment, and RGB colors
- Dynamically added and sized images while maintaining aspect ratios
- Manipulated PowerPoint's built-in placeholders for consistent formatting

### 3. Structured AI Prompt Engineering
- Created detailed prompts that force AI to return structured JSON
- Specified exact formatting rules (newlines, slide types, content structure)
- Used system messages to guide AI behavior consistently
- Implemented fallback content generation when AI responses fail

### 4. Content Generation & Parsing
- Parsed and cleaned AI-generated JSON responses
- Implemented logic to remove markdown formatting from AI responses
- Handled newline-separated content and converted it to PowerPoint bullet points
- Added basic validation for AI responses to ensure proper structure

### 5. Image Processing Workflow
- Retrieved images from Pexels based on search queries
- Saved images temporarily and cleaned up after use
- Created placeholder images when API calls fail
- Implemented logic to maintain image proportions within slide constraints

### 6. Project Architecture
- Created a reusable PPTGenerator class with clear responsibilities
- Used constants for layout settings (margins, ratios, sizes)
- Separated concerns: content generation, slide creation, image handling
- Implemented try-catch blocks for robust error handling

## Key Technical Challenges & Solutions

**AI Response Formatting**: The AI sometimes returned markdown or improperly formatted JSON. Solved by implementing cleaning logic to strip markdown and validate JSON structure.

**Slide Layout Consistency**: Maintaining consistent spacing across different slide types. Solved by creating a configuration system with calculated zones for text and images.

**Image Integration**: Images sometimes distorted or didn't fit slide layout. Solved by implementing maximum height constraints and aspect ratio preservation.

**Content Parsing**: Bullet points weren't displaying as actual bullet points. Solved by using PowerPoint's native bullet formatting instead of manual bullet characters.

This project taught the road to ML is going to be a long one. I am going to walk this path without stopping. I am onto my next project.