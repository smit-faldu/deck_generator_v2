import json
import os
import re
from typing import List, Dict, Any, Optional
from pptx import Presentation
from pptx.dml.color import RGBColor
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain.agents import Tool, AgentExecutor, create_react_agent
from langchain.prompts import PromptTemplate
from langchain.tools import BaseTool
from langchain_core.messages import HumanMessage, AIMessage
from langchain.memory import ConversationBufferMemory
from dotenv import load_dotenv
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.util import Pt

# Load environment variables
load_dotenv()
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")

# Initialize the model
model = ChatGoogleGenerativeAI(
    model="gemini-1.5-pro",
    temperature=0.2,
    google_api_key=GOOGLE_API_KEY,
    convert_system_message_to_human=True
)

class TemplateAnalysisTool(BaseTool):
    """Tool for analyzing PowerPoint templates using the PowerPointAnalyzer"""
    name: str = "analyze_template"
    description: str = "Analyze a PowerPoint template and extract detailed information about its structure"
    memory: Optional[ConversationBufferMemory] = None
    last_analysis: Optional[Dict] = None
    
    def __init__(self, **data):
        super().__init__(**data)
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True,
            output_key="output"
        )
        self.last_analysis = None
    
    def _run(self, template_path: str) -> Dict:
        try:
            # Import the PowerPointAnalyzer
            from powerpoint_analyzer import PowerPointAnalyzer
            
            # Create an analyzer instance
            analyzer = PowerPointAnalyzer(template_path)
            
            # Analyze the template
            analysis = analyzer.analyze()
            
            # Store the analysis in memory
            self.last_analysis = analysis
            
            # Store the interaction in memory
            self.memory.save_context(
                {"input": f"Analyze template: {template_path}"},
                {"output": json.dumps(analysis, default=str)}
            )
            
            return {"analysis": analysis, "status": "success"}
        except Exception as e:
            return {"error": f"Failed to analyze template: {str(e)}", "status": "error"}
    
    def _arun(self, template_path: str):
        raise NotImplementedError("Async version not implemented")

class StructureAnalysisTool(BaseTool):
    """Tool for analyzing PowerPoint structure"""
    name: str = "analyze_structure"
    description: str = "Analyze the PowerPoint structure to determine slide purposes and requirements"
    memory: Optional[ConversationBufferMemory] = None
    last_analysis: Optional[List] = None
    
    def __init__(self, **data):
        super().__init__(**data)
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True,
            output_key="output"
        )
        self.last_analysis = None
    
    def _clean_json_response(self, response_text):
        """Clean up JSON response from the model"""
        # Remove markdown code block markers
        if "```json" in response_text:
            response_text = response_text.split("```json")[1]
        if "```" in response_text:
            response_text = response_text.split("```")[0]
        
        # Remove any leading/trailing whitespace
        response_text = response_text.strip()
        
        # If the response is empty after cleaning, return a default structure
        if not response_text:
            return '{"analysis": []}'
        
        return response_text
    
    def _run(self, template_analysis: Dict) -> Dict:
        try:
            print("\n=== Structure Analysis Process ===")
            print("1. Starting structure analysis...")
            
            # Load template_analysis.json if template_analysis is a string path
            if isinstance(template_analysis, str) and template_analysis.endswith('.json'):
                try:
                    print(f"2. Loading template analysis from file: {template_analysis}")
                    with open(template_analysis, 'r', encoding='utf-8') as f:
                        template_analysis = json.load(f)
                    print("3. Template analysis loaded successfully")
                except Exception as e:
                    print(f"Error loading template analysis file: {str(e)}")
                    return {"error": f"Failed to load template analysis: {str(e)}", "status": "error"}
            
            # Extract slides from the template analysis
            slides = template_analysis.get('slides', [])
            print(f"4. Found {len(slides)} slides to analyze")
            
            # Extract placeholders and shapes from each slide
            structure = []
            for slide in slides:
                slide_info = {
                    'index': slide.get('index', 0),
                    'placeholders': [],
                    'shapes': [],
                    'tables': [],
                    'word_need': [],
                    'layout': slide.get('layout', 'Unknown'),
                    'background': slide.get('background', {}),
                    'notes': slide.get('notes', {})
                }
                
                # Extract placeholders
                for placeholder in slide.get('placeholders', []):
                    text_info = placeholder.get('text', {})
                    if text_info.get('available', False):
                        text = text_info.get('text', '')
                        if text:
                            slide_info['placeholders'].append({
                                'id': placeholder.get('id', ''),
                                'name': placeholder.get('name', ''),
                                'text': text,
                                'position': placeholder.get('position', {}),
                                'type': placeholder.get('type', '')
                            })
                            word_count = len(text.split())
                            slide_info['word_need'].append(word_count)
                
                # Extract shapes
                for shape in slide.get('shapes', []):
                    shape_info = {
                        'id': shape.get('id', ''),
                        'name': shape.get('name', ''),
                        'type': shape.get('type', ''),
                        'position': shape.get('position', {}),
                        'text': {'available': False, 'text': ''},
                        'special_content': shape.get('special_content', {})
                    }
                    
                    # Extract text from shapes
                    text_info = shape.get('text', {})
                    if text_info.get('available', False):
                        text = text_info.get('text', '')
                        if text:
                            shape_info['text'] = {
                                'available': True,
                                'text': text,
                                'paragraphs': text_info.get('paragraphs', [])
                            }
                            slide_info['word_need'].append(len(text.split()))
                    
                    # Extract special content (tables)
                    special_content = shape.get('special_content', {})
                    if special_content.get('type') == 'Table':
                        table_info = {
                            'id': shape.get('id', ''),
                            'name': shape.get('name', ''),
                            'position': shape.get('position', {}),
                            'rows': special_content.get('rows', 0),
                            'columns': special_content.get('columns', 0),
                            'cells': special_content.get('cells', []),
                            'style': {
                                'fill': shape.get('fill', {}),
                                'line': shape.get('line', {})
                            }
                        }
                        slide_info['tables'].append(table_info)
                    
                    slide_info['shapes'].append(shape_info)
                
                if slide_info['placeholders'] or slide_info['shapes'] or slide_info['tables']:
                    structure.append(slide_info)
            
            print("5. Extracted structure from template")
            
            # Generate deep analysis using the model
            prompt = f"""
            Analyze this PowerPoint structure and determine the purpose of each slide.
            For each slide, identify:
            1. The main purpose (e.g., title, introduction, problem statement, etc.)
            2. The type of content needed
            3. The tone and style required
            4. Key elements that should be included
            5. For slides with tables, identify what kind of data should be in each table

            Structure to analyze:
            {json.dumps(structure, indent=2)}

            For each slide, provide the analysis in this format:
            {{
                index: [slide number],
                purpose: [main purpose],
                content: [type of content needed],
                tone: [required tone and style],
                key_elements: [list of key elements],
                word_required: [list of word counts],
                tables: [
                    {{
                        id: [table id],
                        purpose: [purpose of this table],
                        data_type: [type of data that should be in this table],
                        row_headers: [suggested row headers],
                        column_headers: [suggested column headers],
                        style_requirements: {{
                            "header_style": [style for header row],
                            "data_style": [style for data rows],
                            "alignment": [text alignment],
                            "font": [font requirements]
                        }}
                    }}
                ]
            }}
            
            IMPORTANT: Return ONLY the JSON content without any markdown formatting or code block markers.
            """
            
            print("6. Sending prompt to model")
            response = model.invoke(prompt)
            print("7. Received response from model")
            print(f"   Raw response: {response.content[:100]}...")
            
            # Clean up the response
            print("8. Cleaning up model response")
            cleaned_response = self._clean_json_response(response.content)
            print(f"   Cleaned response: {cleaned_response[:100]}...")
            
            # Parse the analysis
            try:
                print("9. Parsing analysis JSON")
                analysis = json.loads(cleaned_response)
                print("10. Analysis parsed successfully")
            except json.JSONDecodeError as e:
                print(f"9. Error parsing analysis: {str(e)}")
                # If parsing fails, create a simple analysis
                analysis = []
                for slide in structure:
                    analysis.append({
                        'index': slide.get('index', 0),
                        'purpose': 'Unknown',
                        'content': 'Generic content',
                        'tone': 'Professional',
                        'key_elements': ['Key point 1', 'Key point 2'],
                        'word_required': slide.get('word_need', [50]),
                        'tables': []
                    })
            
            # Create a structured response that matches template_analysis.json format
            structured_response = {
                "metadata": template_analysis.get("metadata", {}),
                "slides": []
            }
            
            # Add analysis to each slide in the structured response
            for slide in template_analysis.get("slides", []):
                slide_index = slide.get("index", 0)
                
                # Find matching analysis
                slide_analysis = None
                for a in analysis:
                    if a.get('index') == slide_index:
                        slide_analysis = a
                        break
                
                if slide_analysis:
                    # Add analysis to the slide
                    slide["analysis"] = slide_analysis
                
                structured_response["slides"].append(slide)
            
            result = {
                "analysis": analysis, 
                "structure": structure, 
                "structured_response": structured_response,
                "status": "success"
            }
            
            # Store analysis in memory
            self.last_analysis = analysis
            
            # Store the interaction in memory
            self.memory.save_context(
                {"input": "Analyze PowerPoint structure"},
                {"output": json.dumps(analysis)}
            )
            
            # Save the analysis to a JSON file
            try:
                print("11. Saving analysis to JSON file")
                with open('structure_analysis.json', 'w', encoding='utf-8') as f:
                    json.dump(result, f, indent=2, ensure_ascii=False)
                print("12. Analysis saved successfully")
            except Exception as e:
                print(f"Error saving analysis to file: {str(e)}")
            
            print("13. Structure analysis completed successfully")
            return result
        except Exception as e:
            print(f"\nError in structure analysis: {str(e)}")
            print(f"Error type: {type(e)}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            return {"error": str(e), "status": "error"}
    
    def _arun(self, template_analysis: Dict):
        raise NotImplementedError("This tool does not support async")

class ContentGenerationTool(BaseTool):
    """Tool for generating content for PowerPoint slides"""
    name: str = "generate_content"
    description: str = "Generate content for PowerPoint slides based on structure analysis"
    memory: Optional[ConversationBufferMemory] = None
    
    def __init__(self, **data):
        super().__init__(**data)
        self.memory = ConversationBufferMemory(
            memory_key="chat_history",
            return_messages=True,
            output_key="output"
        )
    
    def _clean_json_response(self, response_text):
        """Clean up JSON response from the model"""
        # Remove markdown code block markers
        if "```json" in response_text:
            response_text = response_text.split("```json")[1]
        if "```" in response_text:
            response_text = response_text.split("```")[0]
        
        # Remove any leading/trailing whitespace
        response_text = response_text.strip()
        
        # If the response is empty after cleaning, return a default structure
        if not response_text:
            return '{"content": [{"text": "No content generated", "style": {"font": "Arial", "size": 18, "color": "#000000", "alignment": "left"}}]}'
        
        return response_text
    
    def _get_slide_explanation(self, slide_index: int, template_slide: Dict, slide_analysis: Dict) -> str:
        """Generate a detailed explanation of the slide for the agent"""
        explanation = f"\nSlide {slide_index + 1} Analysis:\n"
        
        # Add template information
        if template_slide:
            explanation += "\nTemplate Structure:\n"
            # Add placeholders
            if template_slide.get('placeholders'):
                explanation += "- Placeholders:\n"
                for ph in template_slide['placeholders']:
                    explanation += f"  * {ph.get('name', 'Unnamed')}: {ph.get('text', {}).get('text', 'No text')}\n"
            
            # Add shapes
            if template_slide.get('shapes'):
                explanation += "- Shapes:\n"
                for shape in template_slide['shapes']:
                    explanation += f"  * {shape.get('name', 'Unnamed')}: {shape.get('type', 'Unknown type')}\n"
                    if shape.get('text', {}).get('available'):
                        explanation += f"    Text: {shape.get('text', {}).get('text', 'No text')}\n"
            
            # Add tables
            if template_slide.get('tables'):
                explanation += "- Tables:\n"
                for table in template_slide['tables']:
                    explanation += f"  * {table.get('name', 'Unnamed')}: {table.get('rows', 0)}x{table.get('columns', 0)} table\n"
        
        # Add analysis information
        if slide_analysis:
            explanation += "\nContent Analysis:\n"
            explanation += f"- Purpose: {slide_analysis.get('purpose', 'Not specified')}\n"
            explanation += f"- Content Type: {slide_analysis.get('content', 'Not specified')}\n"
            explanation += f"- Tone: {slide_analysis.get('tone', 'Not specified')}\n"
            
            if slide_analysis.get('key_elements'):
                explanation += "- Key Elements:\n"
                for element in slide_analysis['key_elements']:
                    explanation += f"  * {element}\n"
            
            if slide_analysis.get('tables'):
                explanation += "- Table Requirements:\n"
                for table in slide_analysis['tables']:
                    explanation += f"  * {table.get('id', 'Unnamed')}: {table.get('purpose', 'Not specified')}\n"
                    explanation += f"    Data Type: {table.get('data_type', 'Not specified')}\n"
        
        return explanation
    
    def _run(self, structure_analysis: Dict, topic: str) -> Dict:
        try:
            print("\n=== Content Generation Process ===")
            print("1. Starting content generation...")
            print(f"Topic: {topic}")
            
            # Load template_analysis.json if it exists
            template_structure = {}
            try:
                print("2. Loading template analysis")
                with open('template_analysis.json', 'r', encoding='utf-8') as f:
                    template_structure = json.load(f)
                print("3. Template analysis loaded successfully")
            except Exception as e:
                print(f"Warning: Could not load template_analysis.json: {str(e)}")
            
            # Extract structured response and analysis
            if isinstance(structure_analysis, dict):
                structured_response = structure_analysis.get('structured_response', {})
                analysis = structure_analysis.get('analysis', [])
                print("4. Found structured response and analysis")
            elif isinstance(structure_analysis, list):
                analysis = structure_analysis
                structured_response = {
                    "metadata": {},
                    "slides": []
                }
                for slide in analysis:
                    structured_response["slides"].append({
                        "index": slide.get('index', 0),
                        "analysis": slide
                    })
                print("4. Found analysis as list, created structured response")
            else:
                print("4. Structure analysis is not a dictionary or list, using empty structures")
                structured_response = {
                    "metadata": {},
                    "slides": []
                }
                analysis = []
            
            print("5. Starting slide content generation")
            # Generate content for each slide
            content = []
            
            # Process each slide in the structured response
            for slide in structured_response.get("slides", []):
                slide_index = slide.get("index", 0)
                slide_analysis = slide.get("analysis", {})
                
                print(f"\nProcessing slide {slide_index + 1}")
                
                # Get template slide structure if available
                template_slide = None
                if template_structure and "slides" in template_structure:
                    for ts in template_structure["slides"]:
                        if ts.get("index") == slide_index:
                            template_slide = ts
                            break
                
                slide_content = {
                    'index': slide_index,
                    'content': [],
                    'word_count': slide_analysis.get('word_required', []),
                    'tables': []
                }
                
                # Generate main content based on slide analysis and template
                print(f"6. Generating main content for slide {slide_index + 1}")
                
                # Get detailed slide explanation
                slide_explanation = self._get_slide_explanation(slide_index, template_slide, slide_analysis)
                
                prompt = f"""
                Generate content for a PowerPoint slide about {topic}.
                
                {slide_explanation}
                
                Generate content that:
                1. EXACTLY matches the template structure shown above
                2. Uses the same formatting, styles, and positions
                3. Follows the content analysis requirements
                4. Maintains the same layout and organization
                5. Uses appropriate AI-focused content
                
                Return the content in this JSON format:
                {{
                    "content": [
                        {{
                            "text": "content text",
                            "style": {{
                                "font": "font name",
                                "size": font size,
                                "color": "color hex",
                                "alignment": "left/center/right"
                            }},
                            "position": {{
                                "left": left position,
                                "top": top position,
                                "width": width,
                                "height": height
                            }},
                            "type": "placeholder/shape/table"
                        }}
                    ]
                }}
                
                IMPORTANT: Return ONLY the JSON content without any markdown formatting or code block markers.
                """
                
                print("7. Sending prompt to model")
                response = model.invoke(prompt)
                print("8. Received response from model")
                print(f"   Raw response: {response.content[:100]}...")
                
                # Clean up the response
                print("9. Cleaning up model response")
                cleaned_response = self._clean_json_response(response.content)
                print(f"   Cleaned response: {cleaned_response[:100]}...")
                
                try:
                    print("10. Parsing content JSON")
                    content_json = json.loads(cleaned_response)
                    slide_content['content'] = content_json.get('content', [])
                    print("11. Content parsed successfully")
                except json.JSONDecodeError as e:
                    print(f"10. Error parsing content: {str(e)}")
                    # If parsing fails, create a simple content entry
                    slide_content['content'] = [{
                        'text': cleaned_response,
                        'style': {
                            'font': 'Arial',
                            'size': 18,
                            'color': '#000000',
                            'alignment': 'left'
                        }
                    }]
                
                # Generate table content if tables are present
                print("12. Checking for tables")
                tables = slide.get('tables', [])
                if tables:
                    print(f"13. Found {len(tables)} tables to process")
                    for table_idx, table in enumerate(tables):
                        print(f"\nProcessing table {table_idx + 1}")
                        
                        # Find matching table analysis
                        table_analysis = None
                        for t in slide_analysis.get('tables', []):
                            if t.get('id') == table.get('id'):
                                table_analysis = t
                                break
                        
                        # Find matching template table
                        template_table = None
                        if template_slide:
                            for tt in template_slide.get('tables', []):
                                if tt.get('id') == table.get('id'):
                                    template_table = tt
                                    break
                        
                        if not table_analysis:
                            print(f"No analysis found for table {table_idx + 1}, using default")
                            table_analysis = {
                                'purpose': 'Data presentation',
                                'data_type': 'Generic data',
                                'row_headers': ['Row 1', 'Row 2'],
                                'column_headers': ['Column 1', 'Column 2'],
                                'style_requirements': {
                                    'header_style': {'font': 'Arial', 'size': 12, 'color': '#000000', 'alignment': 'left'},
                                    'data_style': {'font': 'Arial', 'size': 11, 'color': '#000000', 'alignment': 'left'}
                                }
                            }
                        
                        table_content = {
                            'id': table.get('id', f'table_{table_idx}'),
                            'headers': {
                                'rows': table_analysis.get('row_headers', []),
                                'columns': table_analysis.get('column_headers', [])
                            },
                            'data': [],
                            'style': table_analysis.get('style_requirements', {})
                        }
                        
                        # Add template table information if available
                        if template_table:
                            table_content['template'] = template_table
                        
                        # Generate table data
                        table_prompt = f"""
                        Generate data for a table in slide {slide_index + 1}.
                        
                        Table Template Structure:
                        {json.dumps(template_table if template_table else table, indent=2)}
                        
                        Table Analysis:
                        - Purpose: {table_analysis.get('purpose', 'Not specified')}
                        - Data Type: {table_analysis.get('data_type', 'Not specified')}
                        - Row Headers: {', '.join(table_analysis.get('row_headers', []))}
                        - Column Headers: {', '.join(table_analysis.get('column_headers', []))}
                        
                        Generate data that:
                        1. EXACTLY matches the template table structure
                        2. Uses the same formatting and style
                        3. Maintains the same layout and organization
                        4. Uses appropriate AI-focused content
                        5. Preserves all cell positions and types
                        
                        Return the data in this JSON format:
                        {{
                            "data": [
                                {{
                                    "row": row index,
                                    "column": column index,
                                    "value": "cell value",
                                    "style": {{
                                        "font": "font name",
                                        "size": font size,
                                        "color": "color hex",
                                        "alignment": "left/center/right",
                                        "background": "background color hex"
                                    }},
                                    "position": {{
                                        "left": left position,
                                        "top": top position,
                                        "width": width,
                                        "height": height
                                    }}
                                }}
                            ]
                        }}
                        
                        IMPORTANT: Return ONLY the JSON content without any markdown formatting or code block markers.
                        """
                        
                        print("14. Sending table prompt to model")
                        table_response = model.invoke(table_prompt)
                        print("15. Received table response")
                        print(f"   Raw table response: {table_response.content[:100]}...")
                        
                        # Clean up the table response
                        print("16. Cleaning up table response")
                        cleaned_table_response = self._clean_json_response(table_response.content)
                        print(f"   Cleaned table response: {cleaned_table_response[:100]}...")
                        
                        try:
                            print("17. Parsing table data JSON")
                            table_data = json.loads(cleaned_table_response)
                            table_content['data'] = table_data.get('data', [])
                            print("18. Table data parsed successfully")
                        except json.JSONDecodeError as e:
                            print(f"17. Error parsing table data: {str(e)}")
                            table_content['data'] = []
                        
                        slide_content['tables'].append(table_content)
                
                content.append(slide_content)
            
            # Create a structured content response that matches template_analysis.json format
            structured_content = {
                "metadata": template_structure.get("metadata", structured_response.get("metadata", {})),
                "slides": []
            }
            
            # Add content to each slide in the structured content
            for slide in structured_response.get("slides", []):
                slide_index = slide.get("index", 0)
                
                # Find matching content
                slide_content = None
                for c in content:
                    if c.get('index') == slide_index:
                        slide_content = c
                        break
                
                if slide_content:
                    # Add content to the slide
                    slide["content"] = slide_content
                
                structured_content["slides"].append(slide)
            
            print("\n19. Content generation completed successfully")
            result = {
                "content": content, 
                "structured_content": structured_content,
                "status": "success"
            }
            
            # Store the interaction in memory
            self.memory.save_context(
                {"input": f"Generate content for {topic}"},
                {"output": json.dumps(content)}
            )
            
            # Save the content to a JSON file
            try:
                print("20. Saving content to JSON file")
                with open('content_generation.json', 'w', encoding='utf-8') as f:
                    json.dump(result, f, indent=2, ensure_ascii=False)
                print("21. Content saved successfully")
            except Exception as e:
                print(f"Error saving content to file: {str(e)}")
            
            return result
        except Exception as e:
            print(f"\nError in content generation: {str(e)}")
            print(f"Error type: {type(e)}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            return {"error": str(e), "status": "error"}
    
    def _arun(self, structure_analysis: Dict, topic: str):
        raise NotImplementedError("This tool does not support async")

class PowerPointCreatorTool(BaseTool):
    """Tool for creating PowerPoint presentations"""
    name: str = "create_powerpoint"
    description: str = "Create a PowerPoint presentation from content and template"
    template_path: str = ""
    output_path: str = ""
    
    def __init__(self, **data):
        super().__init__(**data)
        self.template_path = data.get('template_path', '')
        self.output_path = data.get('output_path', '')
    
    def _run(self, content: Dict, structure: Dict) -> Dict:
        try:
            print("\n=== PowerPoint Creation Process ===")
            print("1. Loading template...")
            # Load the template
            prs = Presentation(self.template_path)
            
            # Process each slide
            print("2. Processing slides...")
            
            # Extract structured content
            if isinstance(content, dict):
                structured_content = content.get('structured_content', {})
                print("3. Found structured content")
            else:
                print("3. Content is not a dictionary, using empty structure")
                structured_content = {
                    "metadata": {},
                    "slides": []
                }
            
            # Get slides from the structured content
            slides_content = structured_content.get('slides', [])
            print(f"4. Found {len(slides_content)} slides to process")
            
            # Process each slide
            for slide_content in slides_content:
                slide_index = slide_content.get('index', 0)
                if slide_index >= len(prs.slides):
                    print(f"Warning: Slide index {slide_index} exceeds template slides ({len(prs.slides)})")
                    continue
                
                print(f"5. Processing slide {slide_index + 1}")
                slide = prs.slides[slide_index]
                
                # Get content for this slide
                content_data = slide_content.get('content', {})
                content_items = content_data.get('content', [])
                content_index = 0
                
                print("6. Processing shapes and tables")
                # Process all shapes in the slide
                for shape in slide.shapes:
                    # Handle text shapes
                    if shape.has_text_frame and content_index < len(content_items):
                        content_item = content_items[content_index]
                        new_text = content_item.get('text', '')
                        style = content_item.get('style', {})
                        
                        print(f"   Replacing text with: {new_text[:50]}...")
                        
                        # Store original styles
                        original_styles = []
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                style_info = {
                                    'font_name': run.font.name,
                                    'font_size': run.font.size,
                                    'bold': run.font.bold,
                                    'italic': run.font.italic,
                                    'color': run.font.color.rgb if hasattr(run.font, 'color') and run.font.color is not None else None
                                }
                                original_styles.append(style_info)
                        
                        # Replace text
                        shape.text = new_text
                        
                        # Reapply styles
                        style_index = 0
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if style_index < len(original_styles):
                                    style_info = original_styles[style_index]
                                    run.font.name = style_info['font_name']
                                    run.font.size = style_info['font_size']
                                    run.font.bold = style_info['bold']
                                    run.font.italic = style_info['italic']
                                    if style_info['color'] is not None:
                                        try:
                                            run.font.color.rgb = style_info['color']
                                        except (AttributeError, TypeError):
                                            run.font.color.rgb = RGBColor(0, 0, 0)
                                    style_index += 1
                        
                        content_index += 1
                    
                    # Handle tables
                    if shape.has_table:
                        table = shape.table
                        for row in table.rows:
                            for cell in row.cells:
                                if cell.text and content_index < len(content_items):
                                    content_item = content_items[content_index]
                                    new_text = content_item.get('text', '')
                                    style = content_item.get('style', {})
                                    
                                    # Store original styles
                                    original_styles = []
                                    if cell.text_frame.paragraphs:
                                        for paragraph in cell.text_frame.paragraphs:
                                            for run in paragraph.runs:
                                                style_info = {
                                                    'font_name': run.font.name,
                                                    'font_size': run.font.size,
                                                    'bold': run.font.bold,
                                                    'italic': run.font.italic,
                                                    'color': run.font.color.rgb if hasattr(run.font, 'color') and run.font.color is not None else None
                                                }
                                                original_styles.append(style_info)
                                    
                                    # Replace text
                                    cell.text = new_text
                                    
                                    # Reapply styles
                                    style_index = 0
                                    if cell.text_frame.paragraphs:
                                        for paragraph in cell.text_frame.paragraphs:
                                            for run in paragraph.runs:
                                                if style_index < len(original_styles):
                                                    style_info = original_styles[style_index]
                                                    run.font.name = style_info['font_name']
                                                    run.font.size = style_info['font_size']
                                                    run.font.bold = style_info['bold']
                                                    run.font.italic = style_info['italic']
                                                    if style_info['color'] is not None:
                                                        try:
                                                            run.font.color.rgb = style_info['color']
                                                        except (AttributeError, TypeError):
                                                            run.font.color.rgb = RGBColor(0, 0, 0)
                                                    style_index += 1
                                    
                                    content_index += 1
                
                # Process tables from content
                print("7. Processing additional tables")
                for table_content in content_data.get('tables', []):
                    table_id = table_content.get('id', '')
                    table_data = table_content.get('data', [])
                    table_style = table_content.get('style', {})
                    
                    # Find matching table by ID
                    table_shape = self._find_matching_table(slide, table_id)
                    if table_shape and hasattr(table_shape, 'table'):
                        table = table_shape.table
                        
                        # Add headers
                        headers = table_content.get('headers', {})
                        row_headers = headers.get('rows', [])
                        col_headers = headers.get('columns', [])
                        
                        # Add row headers
                        for i, header in enumerate(row_headers):
                            if i < len(table.rows):
                                cell = table.rows[i].cells[0]
                                cell.text = str(header)
                                self._apply_cell_style(cell, table_style.get('header', {}))
                        
                        # Add column headers
                        for i, header in enumerate(col_headers):
                            if i < len(table.columns):
                                cell = table.rows[0].cells[i + 1]
                                cell.text = str(header)
                                self._apply_cell_style(cell, table_style.get('header', {}))
                        
                        # Add data
                        for data_item in table_data:
                            row = data_item.get('row', 0)
                            col = data_item.get('column', 0)
                            value = data_item.get('value', '')
                            style = data_item.get('style', {})
                            
                            if row < len(table.rows) and col < len(table.columns):
                                cell = table.rows[row].cells[col]
                                cell.text = str(value)
                                self._apply_cell_style(cell, style)
            
            print("8. Preparing to save presentation...")
            # Create a unique filename to avoid conflicts
            import datetime
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            output_dir = os.path.dirname(self.output_path)
            output_filename = os.path.basename(self.output_path)
            name, ext = os.path.splitext(output_filename)
            unique_output_path = os.path.join(output_dir, f"{name}_{timestamp}{ext}")
            
            # Ensure output directory exists
            os.makedirs(output_dir, exist_ok=True)
            
            # Try to save with the unique filename
            try:
                print(f"9. Saving presentation to: {unique_output_path}")
                prs.save(unique_output_path)
                print("10. PowerPoint creation completed successfully")
                return {"status": "success", "output_path": unique_output_path}
            except PermissionError as e:
                # If permission error, try with a different name
                alternative_path = os.path.join(output_dir, f"pitch_deck_{timestamp}{ext}")
                print(f"Permission error, trying alternative path: {alternative_path}")
                prs.save(alternative_path)
                print("10. PowerPoint creation completed successfully with alternative path")
                return {"status": "success", "output_path": alternative_path}
            
        except Exception as e:
            print(f"\nError in PowerPoint creation: {str(e)}")
            print(f"Error type: {type(e)}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            return {"error": str(e), "status": "error"}
    
    def _find_matching_table(self, slide, table_id):
        """Find a table shape that matches the table ID"""
        for shape in slide.shapes:
            if hasattr(shape, 'name') and shape.name == table_id:
                return shape
            if hasattr(shape, 'has_table') and shape.has_table:
                # If no ID match, return the first table found
                return shape
        return None
    
    def _apply_cell_style(self, cell, style):
        """Apply styling to a table cell"""
        # Apply text styling
        for paragraph in cell.text_frame.paragraphs:
            paragraph.alignment = self._get_alignment(style.get('alignment', 'left'))
            
            for run in paragraph.runs:
                font = run.font
                font.name = style.get('font', 'Arial')
                font.size = Pt(style.get('size', 12))
                font.color.rgb = RGBColor(*self._hex_to_rgb(style.get('color', '#000000')))
        
        # Apply cell fill
        if 'background' in style:
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(*self._hex_to_rgb(style['background']))
    
    def _get_alignment(self, alignment):
        """Convert alignment string to PP_ALIGN constant"""
        alignment_map = {
            'left': PP_ALIGN.LEFT,
            'center': PP_ALIGN.CENTER,
            'right': PP_ALIGN.RIGHT
        }
        return alignment_map.get(alignment.lower(), PP_ALIGN.LEFT)
    
    def _hex_to_rgb(self, hex_color):
        """Convert hex color to RGB tuple"""
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    
    def _arun(self, content: Dict, structure: Dict):
        raise NotImplementedError("This tool does not support async")

class PitchDeckAgent:
    """Agent that orchestrates the pitch deck creation process"""
    
    def __init__(self):
        self.tools = [
            TemplateAnalysisTool(),
            StructureAnalysisTool(),
            ContentGenerationTool(),
            PowerPointCreatorTool(template_path="template/Black Elegant and Modern Startup Pitch Deck Presentation.pptx", 
                                output_path="output/Black Elegant and Modern Startup Pitch Deck.pptx")
        ]
        
        self.prompt = PromptTemplate.from_template(
            """You are an expert pitch deck creator. Your task is to create a professional pitch deck using the available tools.
            
            Follow these steps:
            1. Analyze the PowerPoint template using the analyze_template tool
            2. Analyze the structure using the analyze_structure tool to understand slide purposes
            3. Generate appropriate content using the generate_content tool based on the topic and analysis
            4. Create the final PowerPoint presentation using the create_powerpoint tool
            
            When using the tools:
            - For analyze_template: Pass the template path as a string
            - For analyze_structure: Pass the template analysis from the previous step
            - For generate_content: Pass the topic, structure, and analysis from previous steps
            - For create_powerpoint: Pass the template path, content, and output path
            
            If any step fails:
            - Check the error message in the response
            - Try to fix the issue based on the error
            - If the same error occurs three times, stop and report the issue
            
            You have access to these tools:
            {tools}
            
            Use the following format:
            Question: the input question you must answer
            Thought: you should always think about what to do
            Action: the action to take, should be one of [{tool_names}]
            Action Input: the input to the action
            Observation: the result of the action
            ... (this Thought/Action/Action Input/Observation can repeat N times)
            Thought: I now know the final answer
            Final Answer: the final answer to the original input question
            
            Question: {input}
            Thought: {agent_scratchpad}"""
        )
        
        self.agent = create_react_agent(
            llm=model,
            tools=self.tools,
            prompt=self.prompt
        )
        
        self.agent_executor = AgentExecutor(
            agent=self.agent,
            tools=self.tools,
            verbose=True,
            handle_parsing_errors=True,
        )
    
    def create_pitch_deck(self, template_path: str, topic: str, output_path: str = "output/AI_Based_Pitch_Deck.pptx") -> Dict:
        """
        Create a pitch deck using the agent system.
        
        Args:
            template_path (str): Path to the PowerPoint template
            topic (str): Detailed topic information
            output_path (str): Path to save the final PowerPoint
            
        Returns:
            Dict: Result of the pitch deck creation process
        """
        try:
            print("\n=== Starting Pitch Deck Creation Process ===")
            
            # Validate inputs
            if not os.path.exists(template_path):
                return {"status": "error", "error": f"Template file not found: {template_path}"}
            
            # Ensure output directory exists
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # Get tools
            template_tool = self.tools[0]  # TemplateAnalysisTool
            analysis_tool = self.tools[1]  # StructureAnalysisTool
            content_tool = self.tools[2]  # ContentGenerationTool
            powerpoint_tool = self.tools[3]  # PowerPointCreatorTool
            
            print("\n1. Analyzing PowerPoint Template...")
            # Analyze template
            template_result = template_tool._run(template_path)
            if "error" in template_result:
                print(f"Error analyzing template: {template_result['error']}")
                return {"status": "error", "error": template_result["error"]}
            template_analysis = template_result["analysis"]
            print("Template analyzed successfully!")
            
            print("\n2. Analyzing Structure...")
            # Analyze structure
            analysis_result = analysis_tool._run(template_analysis)
            if "error" in analysis_result:
                print(f"Error analyzing structure: {analysis_result['error']}")
                return {"status": "error", "error": analysis_result["error"]}
            structure = analysis_result["structure"]
            analysis = analysis_result["analysis"]
            print("Structure analyzed successfully!")
            
            print("\n3. Generating Content...")
            # Generate content
            content_result = content_tool._run(structure, topic)
            if "error" in content_result:
                print(f"Error generating content: {content_result['error']}")
                return {"status": "error", "error": content_result["error"]}
            content = content_result["content"]
            print("Content generated successfully!")
            
            print("\n4. Creating PowerPoint...")
            # Update the PowerPointCreatorTool with the current paths
            powerpoint_tool.template_path = template_path
            powerpoint_tool.output_path = output_path
            
            # Create PowerPoint
            powerpoint_result = powerpoint_tool._run(content, structure)
            if "error" in powerpoint_result:
                print(f"Error creating PowerPoint: {powerpoint_result['error']}")
                return {"status": "error", "error": powerpoint_result["error"]}
            print("PowerPoint created successfully!")
            
            print("\n=== Pitch Deck Creation Completed ===")
            return {"status": "success", "output_path": output_path}
            
        except Exception as e:
            print(f"\nError in pitch deck creation: {str(e)}")
            return {"status": "error", "error": str(e)}

def main():
    # Example usage
    template_path = "template/Black Elegant and Modern Startup Pitch Deck Presentation.pptx"
    output_path = "output/Black Elegant and Modern Startup Pitch Deck.pptx"
    
    # Detailed topic with company information
    topic = {
        "company": "TechFlow AI",
        "industry": "AI-Powered Business Process Automation",
        "founded": 2023,
        "location": "San Francisco, CA",
        "problem": [
            "Businesses waste 20+ hours per week on repetitive tasks",
            "Manual data entry leads to 15% error rate",
            "Companies lose $50B annually due to inefficient processes",
            "70% of employees report burnout from repetitive work"
        ],
        "solution": [
            "AI-powered workflow automation platform",
            "Reduces manual work by 80%",
            "99.9% accuracy in data processing",
            "Integrates with existing business tools",
            "Customizable for different industries"
        ],
        "target_market": [
            "Mid-size enterprises (100-1000 employees)",
            "Focus on finance, healthcare, and retail sectors",
            "$2B market size in target segments",
            "25% year-over-year market growth"
        ],
        "current_traction": [
            "50+ enterprise clients",
            "$2M ARR",
            "95% customer satisfaction",
            "40% month-over-month growth"
        ],
        "revenue_model": [
            "Subscription-based pricing",
            "Tiered plans: Basic ($499/mo), Pro ($999/mo), Enterprise (Custom)",
            "Additional revenue from customization services",
            "80% gross margin"
        ],
        "funding_need": [
            "Seeking $5M Series A",
            "18-month runway",
            "Product development: 40%",
            "Market expansion: 30%",
            "Team growth: 20%",
            "Operations: 10%"
        ]
    }
    
    # Create pitch deck using the agent system
    agent = PitchDeckAgent()
    result = agent.create_pitch_deck(template_path, json.dumps(topic), output_path)
    
    if result["status"] == "success":
        print(f"\nPitch deck created successfully at: {result['output_path']}")
    else:
        print("\nFailed to create pitch deck")
        print("Error:", result["error"])

if __name__ == "__main__":
    main() 