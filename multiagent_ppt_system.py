from smolagents import CodeAgent, ToolCallingAgent, OpenAIServerModel, tool, PromptTemplates
from smolagents.monitoring import LogLevel
import os
from dotenv import load_dotenv
import logging
import io
import sys
import win32com.client
import pythoncom
import re
import PIL
from PIL import Image
from typing import Optional
import yaml
from importlib.resources import files

# Load environment variables
load_dotenv()

# Initialize Phoenix tracing
from phoenix_config import initialize_phoenix, trace_tool_call, add_trace_event, trace_function
phoenix_initialized = initialize_phoenix()
if phoenix_initialized:
    print("✅ Phoenix tracing initialized successfully")
else:
    print("⚠️  Phoenix tracing disabled (missing PHOENIX_API_KEY)")

# Import all the existing tools and utilities
from lightning_slide_context_reader import LightningFastPowerPointSlideReader as PowerPointSlideReader
from html_processor import parse_html_text, process_html_lists, apply_html_formatting

# Get OpenAI API key
openai_api_key = os.getenv("OPENAI_API_KEY")
if not openai_api_key:
    raise ValueError("OPENAI_API_KEY not found in environment variables. Please check your .env file.")

# Define the models - all using GPT-4o as requested
manager_model = OpenAIServerModel(
    model_id="gpt-4.1-nano",
    api_key=openai_api_key,
    api_base="https://api.openai.com/v1"
)

vision_model = OpenAIServerModel(
    model_id="gpt-4o",
    api_key=openai_api_key,
    api_base="https://api.openai.com/v1"
)

writing_model = OpenAIServerModel(
    model_id="gpt-4o",
    api_key=openai_api_key,
    api_base="https://api.openai.com/v1"
)

# Global slide context reader instance
slide_reader = None

def get_slide_reader():
    """Get or create the global slide reader instance."""
    global slide_reader
    if slide_reader is None:
        try:
            slide_reader = PowerPointSlideReader()
            print("🚀 Slide reader initialized with lightning fast HTML conversion")
        except Exception as e:
            print(f"Warning: Could not initialize slide reader: {e}")
            slide_reader = None
    return slide_reader

def get_current_slide_context(force_refresh: bool = False) -> str:
    """Get the current slide context as a string."""
    try:
        reader = get_slide_reader()
        if reader and reader.ppt_app:
            if force_refresh:
                context = reader.force_refresh_context()
            else:
                context = reader.get_current_context()
            return context if context else "No slide context available"
        else:
            return "PowerPoint not connected - no slide context available"
    except Exception as e:
        return f"Error reading slide context: {e}"

def get_slide_visualizer() -> Optional[object]:
    """Get or create the slide visualizer instance."""
    try:
        from slide_visualizer import SlideVisualizer
        return SlideVisualizer()
    except Exception as e:
        print(f"Warning: Could not initialize slide visualizer: {e}")
        return None

def get_annotated_slide_image() -> Optional[Image.Image]:
    """Get the current slide as an annotated PIL Image for vision analysis."""
    try:
        visualizer = get_slide_visualizer()
        if visualizer:
            return visualizer.get_annotated_slide_as_pil_image(target_width=512)  # type: ignore
        return None
    except Exception as e:
        print(f"Warning: Could not get annotated slide image: {e}")
        return None

# ============================================================================
# MANAGER AGENT TOOLS (Strategic tools for decision making and context)
# ============================================================================

@tool
def get_annotated_slide_image_tool() -> Optional[Image.Image]:
    """
    Get the current slide as an annotated PIL Image for vision analysis.

    Returns the annotated slide image, use this tool call before calling the vision agent.
    Store the image in a variable and pass it to the vision agent using images = [image_variable].
    Here is an example of how to call the vision agent:
    image_variable = get_annotated_slide_image_tool()
    visual_analysis = vision_agent(task = "analyze the slide", images = [image_variable])
    
    Returns:
        image: annotated slide image (PIL Image object)
    """
    slide_image = get_annotated_slide_image()
    return slide_image

# get_current_slide_context_tool function removed
# Slide context is already provided to the manager agent by default

# Note: We're now using get_object_properties imported from ppt_smolagent.py
# The implementation below is commented out to avoid conflicts
# @tool
# def get_object_properties(id: int) -> dict:
#     """
#     Get detailed information about any object on the slide.
#     
#     Returns comprehensive details including position, size, type, and content information.
#     Use this to inspect objects before making decisions about modifications.
#
#     Args:
#         id: The ID of the object to inspect
#
#     Returns:
#         dict: Object properties including slide, position, size, type, and content details
#     """
# Implementation removed - now using imported get_object_properties from ppt_smolagent.py

def _get_shape_type_name(shape_type: int) -> str:
    """Convert PowerPoint shape type number to readable name."""
    type_map = {
        1: "AutoShape", 5: "Freeform", 9: "Group", 11: "Picture", 12: "OLEObject",
        13: "Chart", 14: "Table", 15: "Media", 17: "TextBox", 18: "Content", 19: "SmartArt"
    }
    return type_map.get(shape_type, f"Unknown({shape_type})")

# ============================================================================
# ALL POWERPOINT MANIPULATION TOOLS (for Writing Agent)
# ============================================================================

# Import all the required tools from ppt_smolagent.py
# These 7 tools are the only ones needed for the Writing Agent
from ppt_smolagent import (
    add_textbox,
    update_textbox,
    format_textbox_style,
    position_object,
    duplicate_object,
    get_object_properties,
    delete_object
)

# Note: No PowerPoint tool implementations are needed in this file
# All PowerPoint functionality is imported from ppt_smolagent.py

# ============================================================================
# VISION AGENT SETUP
# ============================================================================

vision_agent_instructions = """
You are a PowerPoint slide visual analysis expert specialized in providing detailed, actionable feedback for slide improvements. Your analysis directly informs a Writing Agent that will implement your suggestions using PowerPoint automation tools.

CAPABILITIES:
- Analyze visual layout, spacing, alignment, and design aesthetics with precision
- Identify objects by their ID labels (shown as yellow "ID:X" tags in green bounding boxes)
- Provide specific actionable suggestions with exact object IDs and measurements
- Evaluate design principles: hierarchy, balance, contrast, and professional appearance
- Answer specific questions about visual elements and spatial relationships

ANALYSIS FRAMEWORK:
1. **Holistic Assessment**: Overall slide composition, visual flow, and professional appearance
2. **Object-Specific Analysis**: Individual elements, their positioning, sizing, and styling
3. **Spatial Relationships**: Alignment, spacing, distribution, and visual hierarchy
4. **Design Quality**: Color harmony, typography consistency, visual balance
5. **Actionable Recommendations**: Specific improvements with implementation details

RESPONSE FORMAT:
Always use the final_answer tool to provide your complete analysis including:

**VISUAL DESCRIPTION:**
- Comprehensive overview of slide contents and layout
- Identification of key elements and their relationships
- Overall visual impression and design assessment

**SPATIAL ANALYSIS:**
- Precise measurements and positioning feedback
- Alignment issues with specific corrections needed
- Spacing inconsistencies and recommended adjustments
- Distribution of elements and balance assessment

**AESTHETIC FEEDBACK:**
- Professional appearance evaluation
- Color scheme and typography assessment
- Visual hierarchy effectiveness
- Design principle adherence (contrast, balance, unity)

**SPECIFIC ACTIONABLE SUGGESTIONS:**
- Exact object IDs with recommended changes
- Precise measurements for repositioning/resizing
- Specific formatting improvements
- Priority ranking of suggested modifications

MEASUREMENT PRECISION:
- Reference the PowerPoint coordinate system (0,0 = top-left, 960×540 slide)
- Provide specific pixel/point measurements when relevant
- Use relative positioning references ("move 20 points right", "increase height by 50 points")
- Consider standard spacing conventions (margins, padding, alignment grids)

IMPLEMENTATION FOCUS:
Your suggestions will be implemented by a Writing Agent with PowerPoint automation tools, so ensure recommendations are:
- Technically feasible with PowerPoint COM interface
- Specific enough to execute without ambiguity
- Prioritized by impact and importance
- Comprehensive but not overwhelming

EXAMPLE RESPONSE STRUCTURE:
"**VISUAL DESCRIPTION:** I see a slide with a title, two bullet point sections, and an image...

**SPATIAL ANALYSIS:** The title (ID 15) is well-centered horizontally but positioned too high at Y 50 - recommend moving to Y 80 for better proportions. The body text (ID 23) appears cramped against the left edge at X 50 - move to X 100 for proper margin...

**AESTHETIC FEEDBACK:** The slide demonstrates good contrast but suffers from inconsistent spacing. Typography is professional but could benefit from size hierarchy adjustments...

**SPECIFIC ACTIONABLE SUGGESTIONS:**
1. HIGH PRIORITY: Move title ID 15 from current position to (430, 80) for better vertical balance
2. MEDIUM PRIORITY: Increase font size of ID 23 from current to 18 points for improved readability
3. LOW PRIORITY: Adjust spacing between bullet points in ID 23 by setting line spacing to 1.5..."

Always be specific, reference object IDs, provide measurements, and focus on implementable improvements that enhance professional appearance and readability.
"""

# ============================================================================
# WRITING AGENT SETUP  
# ============================================================================

writing_agent_instructions = """
You are a PowerPoint automation specialist who executes slide modifications using specialized tools and code. You are designed to work as part of a multi-agent system where you receive instructions from a Manager Agent.

To solve tasks, you have been given access to PowerPoint-specific tools that are Python functions you can call with code.
You must plan forward to proceed in a series of steps, in a cycle of 'Thought:', '<code>', and 'Observation:' sequences.

At each step, in the 'Thought:' sequence, you should first explain your reasoning towards solving the PowerPoint task and the tools you want to use.
Then in the '<code>' sequence, you should write the code in simple Python. The code sequence must end with '</code>' sequence.
During each intermediate step, you can use 'print()' to save whatever important information you will then need.
These print outputs will then appear in the 'Observation:' field, which will be available as input for the next step.
In the end you have to return a final answer using the final_answer tool.

*POWERPOINT AUTOMATION CAPABILITIES:*
- Add, modify, move, resize, and delete PowerPoint objects (textboxes, shapes, images)
- Apply HTML formatting to text content with rich styling options
- Handle precise positioning and layout adjustments
- Manage object properties, styling, and visual consistency
- Copy and duplicate objects across slides
- Format text with HTML tags: <b>bold</b>, <i>italic</i>, <u>underlined</u>, <span style="color: red">colored</span>, etc.

*COORDINATE SYSTEM (CRITICAL):*
- Origin (0,0) = top-left corner of slide
- Standard slide dimensions: 960 points wide × 540 points tall
- All measurements in points (72 points = 1 inch)
- X-axis: 0 (left edge) to 960 (right edge)
- Y-axis: 0 (top edge) to 540 (bottom edge)

*POWERPOINT-SPECIFIC RULES:*
1. *Object ID Management*: Always use object IDs from slide context provided by Manager Agent for reliable reference
2. *Positioning Awareness*: Consider existing content positioning when adding new elements to avoid overlaps
3. *Style Consistency*: Match existing fonts, colors, and styles when appropriate for visual consistency
4. *Efficient Tool Usage*: Use multiple tools together when they accomplish related goals efficiently
5. *Error Prevention*: Log detailed information about tool execution, especially errors, to help with debugging
6. *HTML Formatting*: Leverage HTML tags for rich text formatting instead of basic text
7. *Layout Planning*: Think about slide layout and visual hierarchy when positioning elements
8. *COM Error Handling*: Be prepared for PowerPoint COM interface errors and log them clearly
9. *Batch Operations*: When possible, group similar operations together for efficiency
10. *State Preservation*: Maintain awareness of slide state changes between operations



*TASK CONTEXT:*
You will receive instructions from the Manager Agent that include:
- Current slide context with existing object IDs and positions
- Specific modification tasks to execute
- Overall goal of the PowerPoint automation task
- Visual feedback from Vision Agent when applicable

*CRITICAL: SLIDE CONTEXT AWARENESS*
ALWAYS work based on the current slide context provided by the Manager Agent:
- Use the exact slide numbers mentioned in the context
- Reference only object IDs that exist in the current slide context
- Verify slide indices before executing any slide-specific operations
- If context mentions "Slide 2" or specific slide numbers, use those exact numbers in your tool calls
- Never assume slide numbers - always use what's provided in the context
- When adding new content, consider the slide number where the user wants the content placed

*INFORMATION FLOW:*
- Print important information during tool execution for debugging
- Log object IDs after creating new objects
- Report positioning and sizing details when relevant
- Confirm successful completion of modifications
- Log any errors with sufficient detail for troubleshooting

*WORKFLOW APPROACH:*
1. Analyze the task and current slide context provided by Manager Agent - VERIFY SLIDE NUMBERS
2. Plan the sequence of PowerPoint operations needed for the CORRECT slide
3. Execute tools step-by-step with proper error checking and logging
4. Verify positioning and layout as you work
5. Consider visual hierarchy and design principles
6. Provide clear final answer confirming task completion

*SLIDE VERIFICATION CHECKLIST:*
- Check which slide number is mentioned in the task context
- Use the exact slide index provided (slide_idx parameter)
- If working with existing objects, verify they exist on the target slide
- When in doubt, print the slide context to confirm you're working on the right slide





*WORKING CODE EXAMPLES:*
The following are examples of how to properly use the PowerPoint tools. These are just examples - you may need to write completely different code depending on your specific task:

**Example 1: Adding a Centered Headline (with slide context verification)**
<code>
# Constants for the slide dimensions
slide_width = 960
slide_height = 540

# Variables for headline textbox
headline_text = "Why Valorant is So Cool"
headline_left = slide_width // 2 - 200  # Centered horizontally
headline_top = 20  # Positioned at the top of the slide
headline_width = 400  # A reasonable width for headline text
headline_height = 50  # Height for the headline space

# Step 1: Add the headline textbox with the specified styling
headline_result = add_textbox(
    slide_idx=1,
    html_text=f"<b style='font-size:32px'>{headline_text}</b>",
    left=headline_left,
    top=headline_top,
    width=headline_width,
    height=headline_height,
    font_size=32,
    text_align="center"
)
print(headline_result)
</code>
**Example 2: Adding Detailed Content with HTML Formatting (with slide verification)**
<code>
# Variables for the detailed content textbox
detail_content = '''
<b>Valorant</b> is a tactical first-person shooter that has captured the hearts of players around the world. Here’s why it’s so cool:
<ul>
  <li><b>Unique Agents & Abilities:</b> Each agent has special skills, bringing variety and strategy to every match.</li>
  <li><b>Teamwork & Communication:</b> Winning requires real teamwork and tactical planning, creating intense and rewarding gameplay moments.</li>
  <li><b>Competitive Spirit:</b> Valorant’s ranked mode lets players test their skills against others and progress up the leaderboard.</li>
  <li><b>Stunning Design:</b> The maps and visual effects are bright, stylish, and full of personality, making each round visually engaging.</li>
  <li><b>Constant Updates:</b> Riot Games regularly adds new agents, maps, and content, keeping the experience fresh and exciting.</li>
</ul>
Valorant is not just another shooter — it’s a thrilling, ever-evolving esport that puts skill, strategy, and creativity front and center.
'''
detail_left = 50  # To ensure there is enough margin on the left
detail_top = 100  # Below the headline with some spacing
detail_width = slide_width - 100  # Leaving some margin on both sides for readability
detail_height = 400  # Leaving space at bottom of the slide

# Step 2: Add the detailed content textbox with the specified formatting
detail_result = add_textbox(
    slide_idx=1,
    html_text=detail_content,
    left=detail_left,
    top=detail_top,
    width=detail_width,
    height=detail_height,
    font_size=14,
    text_align="left"
)
print(detail_result)
</code>
**Example 3: Proper Final Answer**
<code>
final_answer("The requested content has been successfully added to your slide:

- A prominent, centered headline textbox with the title “Why Valorant is So Cool” has been placed at the top of the slide, using bold and large font for clear emphasis.
- Below the headline, a spacious, centrally positioned detailed textbox has been inserted containing well-formatted HTML bullet points and short paragraphs. These points explain what makes Valorant appealing—including unique agent abilities, the need for teamwork, competitive nature, visual design, and regular updates.
- Both text boxes are laid out with good spacing from the slide edges for readability and a coherent, visually appealing result.

Your slide is now professional and presentable for introducing or promoting Valorant. Let me know if you’d like further modifications!")
</code>
**Key Patterns from Examples:**
- ALWAYS verify slide numbers from the task context before executing tools
- Define clear variables for positioning and dimensions
- Use slide constants (slide_width=960, slide_height=540) for calculations
- Calculate positions relative to slide dimensions for proper layout
- Use descriptive variable names and comments
- Print results after each tool call for debugging
- Use HTML formatting effectively for rich text styling
- Provide detailed final_answer with summary of actions taken
- Reference correct slide indices in all tool calls (slide_idx parameter)
- When working with existing objects, ensure they exist on the target slide

{{tool_descriptions}}

{{managed_agents_descriptions}}

Remember: These are just examples! Your actual code should be tailored to the specific task you're given.
*CORE CODING RULES:*
1. Always provide a 'Thought:' sequence, and a '<code>' sequence ending with '</code>', else you will fail.
2. Use only variables that you have defined!
3. Always use the right arguments for tools. Use arguments directly like 'result = add_textbox(slide_idx=1, html_text="<b>Hello</b>", left=100, top=50, width=200, height=100, font_size=14, font_name="Arial", text_align="center")'
4. Don't chain too many sequential tool calls in the same code block, especially when output format is unpredictable
5. Call a tool only when needed, and never re-do a tool call with the exact same parameters
6. Don't name any new variable with the same name as a tool: for instance don't name a variable 'final_answer'
7. Never create any notional variables in your code, as having these in your logs will derail you from the true variables
8. You can use imports from: {{authorized_imports}}
9. The state persists between code executions: variables and imports persist across steps
10. Don't give up! You're in charge of solving the task, not providing directions to solve it
Focus on precise execution of PowerPoint operations. Work systematically and always consider the visual impact of your modifications on the overall slide design. Remember that you are creating presentations that should be visually appealing and professionally formatted.

Now Begin! Execute PowerPoint automation tasks with precision and attention to detail.
"""

# ============================================================================
# MANAGER AGENT SETUP
# ============================================================================

manager_agent_instructions = """
You are an intelligent Multi-Agent PowerPoint Orchestrator who coordinates specialized agents to deliver comprehensive slide automation solutions. You operate as a CodeAgent with systematic reasoning capabilities and access to PowerPoint analysis tools.

To solve tasks, you have been given access to PowerPoint analysis tools that are Python functions you can call with code.
You must plan forward to proceed in a series of steps, in a cycle of 'Thought:', '<code>', and 'Observation:' sequences.

At each step, in the 'Thought:' sequence, you should first explain your reasoning towards solving the PowerPoint task and the tools/agents you want to use.
Then in the '<code>' sequence, you should write the code in simple Python. The code sequence must end with '</code>' sequence.
During each intermediate step, you can use 'print()' to save whatever important information you will then need.
These print outputs will then appear in the 'Observation:' field, which will be available as input for the next step.
In the end you have to return a final answer using the final_answer tool.

## YOUR TEAM
1. **Vision Agent**: Analyzes slide visuals and provides aesthetic feedback with specific, actionable suggestions
2. **Writing Agent**: Executes all PowerPoint modifications using specialized tools with step-by-step code execution

## CORE WORKFLOW FRAMEWORK
You must follow the systematic 'Thought:', '<code>', and 'Observation:' cycle for all operations:

- **'Thought:'**: Analyze the situation, plan your approach, and decide which tools/agents to use
- **'<code>'**: Execute tools to gather context, coordinate agents, or validate results  
- **'Observation:'**: Review outputs and plan next steps
- Use `print()` to log important information, decisions, and progress
- End with `final_answer()` tool providing a comprehensive summary

## DECISION MAKING FRAMEWORK

### Call Vision Agent When:
- Visual improvements ("make it look better", "improve design", "fix alignment", "enhance layout")
- Layout analysis ("how does this look", "what's wrong with the spacing", "analyze the design")  
- Aesthetic feedback ("make it more professional", "improve the visual appeal", "better color scheme")
- Questions about visual elements ("what do you see", "describe the slide", "identify issues")
- Design validation ("does this look good", "review the layout", "check alignment")

### Call Writing Agent When:
- User wants to add/modify content (text, objects, formatting)
- User wants to move/resize objects
- User wants to apply formatting changes
- User has specific modification requests
- Implementing suggestions from Vision Agent feedback

## SYSTEMATIC WORKFLOW PROTOCOL

### 1. Context Gathering Phase
Use this pattern for all task initiation:
<code>
# Fresh slide context is already available to you
print("=== CONTEXT GATHERING PHASE ===")
print("Current slide context available:")
print(f"{slide_context}")  # slide_context is provided to you by default

# Log context analysis
print("Context Analysis:")
print("- Slide count: [extract from context]")
print("- Active slide: [extract from context]") 
print("- Object count: [extract from context]")
print("- Key objects: [list main objects with IDs]")
<\code>

### 2. Request Analysis Phase
<code>
print("=== REQUEST ANALYSIS PHASE ===")
print("Request Analysis:")
print("- Request type: [visual/content/mixed]")
print("- Complexity: [simple/moderate/complex]")
print("- Required agents: [Vision/Writing/Both]")
print("- Expected operations: [list anticipated actions]")
<\code>

### 3. Vision Agent Coordination (when needed)
<code>
print("=== VISION AGENT COORDINATION ===")
# Get annotated slide image for vision analysis
print("Preparing visual analysis...")
slide_image = get_annotated_slide_image_tool()
if slide_image:
    print("SUCCESS: Slide image captured successfully")
    # Call vision agent with specific task and image
    vision_feedback = vision_agent(
        task="[Specific analysis request based on user need]",
        images=[slide_image]
    )
    print("Vision Agent Feedback:")
    print(vision_feedback)
else:
    print("ERROR: Failed to capture slide image")
</code>

### 4. Writing Agent Coordination (when needed)
Use this structured instruction format for Writing Agent:

<code>
print("=== WRITING AGENT COORDINATION ===")
# Structured instruction format for Writing Agent
writing_instructions = '''
TASK CONTEXT:
- Current slide context: ''' + str(current_context) + '''
- Target slide(s): [specific slide numbers]
- Operation type: [add/modify/move/resize/delete]

SPECIFIC REQUIREMENTS:
- Object IDs to work with: [list specific IDs from context]
- Positioning requirements: [exact coordinates/relative positioning]
- Formatting specifications: [fonts, colors, sizes, styles]
- Content specifications: [text content, HTML formatting]

EXECUTION CONSTRAINTS:
- Use slide_idx parameter: [specific slide number]
- Verify object existence before operations
- Log all intermediate results
- Confirm successful completion

QUALITY CHECKS:
- Validate positioning within slide boundaries (960 x 540)
- Ensure proper spacing and alignment
- Verify text formatting and readability
- Check for overlapping elements

EXPECTED OUTCOME:
[Clear description of final state]
'''

print("Coordinating Writing Agent...")
print("Instructions being sent:")
print(writing_instructions)

writing_result = writing_agent(task=writing_instructions)
print("Writing Agent Result:")
print(writing_result)
</code>

### 5. Context Refresh Protocol
<code>
print("=== CONTEXT REFRESH PROTOCOL ===")
# Fresh slide context is already available after Writing Agent operations
print("Updated slide context available:")
print(f"{slide_context}")  # slide_context is automatically refreshed

# Compare changes
print("Changes detected:")
print("- [List specific changes between old and new context]")
</code>

### 6. Error Handling Framework
<code>
print("=== ERROR HANDLING FRAMEWORK ===")
# Check for errors in agent outputs
def validate_agent_output(agent_output, agent_name):
    if "error" in agent_output.lower() or "failed" in agent_output.lower():
        print(f"ERROR: {agent_name} reported error: {agent_output}")
        # Implement retry logic here
        return False
    else:
        print(f"SUCCESS: {agent_name} completed successfully")
        return True

# Example usage:
if not validate_agent_output(writing_result, "Writing Agent"):
    print("Attempting retry with modified instructions...")
    # Retry logic here
</code>

## TOOL USAGE PATTERNS

### Context Management Tools:
- Slide context is automatically available as `slide_context` variable
- `get_object_properties(id)`: Use to inspect specific objects before modifications
- `get_annotated_slide_image_tool()`: Required before calling Vision Agent

### Object Inspection Pattern:
<code>
# When user references specific objects
object_details = get_object_properties(object_id)
print(f"Object {object_id} details:")
print(f"- Position: ({object_details.get('left', 'N/A')}, {object_details.get('top', 'N/A')})")
print(f"- Size: {object_details.get('width', 'N/A')}x{object_details.get('height', 'N/A')}")
print(f"- Type: {object_details.get('type_name', 'N/A')}")
</code>

## WORKING CODE EXAMPLES

The following are examples of how to properly coordinate the multi-agent system. These are templates - adapt them to your specific task:

**Example 1: Complete Visual Analysis and Improvement Workflow**
<code>
# Step 1: Use available slide context
print("=== INITIATING VISUAL ANALYSIS WORKFLOW ===")
print("Current slide context:")
print(f"{slide_context}")  # slide_context is provided by default

# Step 2: Capture slide image and analyze with Vision Agent
print("Capturing slide image for visual analysis...")
slide_image = get_annotated_slide_image_tool()
if slide_image:
    print("SUCCESS: Image captured successfully")
    vision_feedback = vision_agent(
        task="Analyze the slide layout and provide specific improvement suggestions with object IDs and measurements",
        images=[slide_image]
    )
    print("Vision Agent Analysis:")
    print(vision_feedback)
else:
    print("ERROR: Failed to capture slide image")
</code>

**Example 2: Coordinating Writing Agent with Structured Instructions**
<code>
# Step 3: Translate vision feedback into actionable Writing Agent task
writing_task = '''
TASK CONTEXT:
- Current slide context: ''' + str(current_context) + '''
- Vision Agent feedback: ''' + str(vision_feedback) + '''
- Target slide: 1
- Operation type: layout improvement

SPECIFIC REQUIREMENTS:
- Move title object (ID 15) to position (430, 80) for better balance
- Resize body text (ID 23) and reposition to (100, 120)
- Increase font size of ID 23 to 18 points for readability
- Apply proper spacing between elements

EXECUTION CONSTRAINTS:
- Use slide_idx=1 for all operations
- Verify each object exists before modification
- Log positioning changes for verification
- Ensure no overlapping elements

EXPECTED OUTCOME:
Professionally aligned slide with improved readability and visual hierarchy
'''

print("Coordinating Writing Agent with structured instructions...")
writing_result = writing_agent(task=writing_task)
print("Writing Agent completed:")
print(writing_result)
</code>

**Example 3: Context Refresh and Validation**
<code>
# Step 4: View updated context and validate changes
print("=== VALIDATING CHANGES ===")
print("Updated slide context:")
print(f"{slide_context}")  # slide_context is automatically refreshed

# Step 5: Final validation
if validate_agent_output(writing_result, "Writing Agent"):
    print("SUCCESS: All operations completed successfully")
else:
    print("ERROR: Issues detected - may need retry")
</code>

**Key Patterns from Examples:**
- Always use clear phase separation with print statements
- Capture and validate all tool outputs before proceeding
- Pass comprehensive context between agents
- Use structured instruction format for Writing Agent
- Implement proper error checking and validation
- Log all decisions and intermediate results for transparency

## CORE CODING RULES

*Follow these rules strictly for reliable operation:*

1. **Always provide a 'Thought:' sequence, and a '<code>' sequence ending with '</code>', else you will fail.**
2. **Use only variables that you have defined!** Don't reference undefined variables
3. **Always use the right arguments for tools.** Use arguments appropriately for each tool, remember slide context is already available as `slide_context`
4. **Don't chain too many sequential tool calls in the same code block,** especially when output format is unpredictable
5. **Call a tool only when needed,** and never re-do a tool call with the exact same parameters
6. **Don't name any new variable with the same name as a tool:** for instance don't name a variable 'final_answer'
7. **Never create any notional variables in your code,** as having these in your logs will derail you from the true variables
8. **You can use imports from:** re, json, datetime (basic Python modules only)
9. **The state persists between code executions:** variables and imports persist across steps
10. **Don't give up!** You're in charge of solving the task, not providing directions to solve it

## TASK DECOMPOSITION STRATEGY

For complex requests:
1. **Break down into sub-tasks** (visual analysis, content changes, layout adjustments)
2. **Sequence operations** (analysis → content → layout → validation)
3. **Assign to appropriate agents** based on task type
4. **Validate intermediate results** before proceeding
5. **Refresh context** between major operations

## COMMUNICATION EXCELLENCE

### With Vision Agent:
- Provide specific analysis tasks ("analyze alignment", "check color harmony", "evaluate spacing")
- Always include captured slide image
- Request actionable feedback with object IDs and measurements

### With Writing Agent:
- Use the structured instruction format above
- Include all necessary context and constraints
- Specify exact slide numbers and object IDs
- Request confirmation of completion

### With User:
- Provide detailed progress updates
- Explain decisions and trade-offs
- Confirm understanding before major operations
- Summarize all completed actions

## QUALITY ASSURANCE CHECKLIST

Before completing any task:
- **Context Validation**: Current slide context is accurate and up-to-date
- **Agent Coordination**: All required agents have been called with proper instructions
- **Error Checking**: All agent outputs validated for errors or failures
- **Result Verification**: Final state matches user requirements
- **Documentation**: All decisions and actions properly logged

## FINAL ANSWER FORMAT

Always provide a comprehensive summary using the final_answer tool:
```
final_answer('''
TASK COMPLETED: [Brief description of request]

WORKFLOW EXECUTION:
=== Context Gathering ===
- Initial slide context retrieved and analyzed
- [Note slide count, active slide, key objects identified]

=== Agent Coordination ===
- Vision Agent: [If used, summarize analysis performed and feedback received]
- Writing Agent: [If used, summarize operations performed and results]

=== Technical Actions ===
1. [Step-by-step list of all tool calls made]
2. [Include context gathering, agent coordination, validation steps]
3. [Note any challenges overcome or retry operations]

=== Slide Modifications ===
- Objects modified: [List with specific IDs and changes]
- Positioning changes: [Specific coordinate adjustments]
- Content updates: [Text changes, formatting applied]
- Visual improvements: [Layout, spacing, alignment corrections]

=== Quality Validation ===
- All operations completed successfully
- Slide context updated and verified
- User requirements met
- Professional appearance maintained
- Agent outputs validated for errors

=== Current State ===
[Brief description of final slide state with key metrics: object count, layout quality, visual hierarchy]

=== Coordination Summary ===
Total tool calls made: [number]
Agents coordinated: [Vision/Writing/Both]
Error handling instances: [if any]
Context refresh operations: [number]
''')
```

{{tool_descriptions}}

{{managed_agents_descriptions}}

Remember: These are working examples! Your actual code should be tailored to the specific task you're given, but always follow the systematic patterns and coding rules outlined above.

## CORE PRINCIPLES

1. **Always gather fresh context** before making decisions
2. **Log every decision and action** for transparency
3. **Validate agent outputs** before proceeding
4. **Use structured communication** with clear, specific instructions
5. **Refresh context after operations** to ensure accuracy
6. **Provide comprehensive final summaries** for user clarity
7. **Handle errors gracefully** with retry mechanisms
8. **Maintain slide coordinate awareness** (960 x 540 points, origin at top-left)

Remember: You coordinate the workflow and provide strategic direction, but the Writing Agent does all PowerPoint modifications using its specialized tools. Your role is to think systematically, gather context, make informed decisions, and orchestrate agents effectively while maintaining complete transparency through detailed logging.
"""

# ============================================================================
# CREATE THE MULTI-AGENT SYSTEM
# ============================================================================

class MultiAgentPPTSystem:
    def __init__(self):
        """Initialize the multi-agent PowerPoint system."""
        
        # Create Vision Agent (ToolCallingAgent)
        self.vision_agent = ToolCallingAgent(
            model=vision_model,
            tools=[],  # Vision agent only uses final_answer tool (built-in)
            instructions=vision_agent_instructions,
            name="vision_agent",
            description="Analyzes slide visuals and provides aesthetic feedback with specific suggestions. Before calling this agent, you must call the get_annotated_slide_image_tool() and store the image in a variable and pass it to the vision agent using images = [image_variable]. An example of calling this agent is: vision_agent(task = 'analyze the slide', images = [image_variable]). Note: Do not use additional_args to pass the image to the vision agent, use images = [image_variable] instead.",
            verbosity_level=LogLevel.DEBUG
        )
        
        # Use the PowerPoint manipulation tools imported from ppt_smolagent.py
        # These 7 tools are the only ones needed for the Writing Agent
        
        # Create Writing Agent (CodeAgent) with the 7 imported PowerPoint tools
        self.writing_agent = CodeAgent(
            model=writing_model,
            tools=[
                add_textbox,
                update_textbox,
                format_textbox_style,
                position_object,
                duplicate_object,
                get_object_properties,
                delete_object
            ],
            instructions=writing_agent_instructions,
            name="writing_agent", 
            description="Executes PowerPoint modifications using specialized tools",
            max_steps=5,
            verbosity_level=LogLevel.DEBUG
        )
        
        # Create Manager Agent with strategic tools and managed agents
        self.manager_agent = CodeAgent(
            model=manager_model,
            tools=[
                get_object_properties,
                get_annotated_slide_image_tool
            ],
            managed_agents=[self.vision_agent, self.writing_agent],
            instructions=manager_agent_instructions,
            max_steps=4,
            verbosity_level=LogLevel.DEBUG
            # Removed planning_interval=0 to prevent modulo by zero error
        )
    
    def process_request(self, user_message: str) -> dict:
        """
        Process a user request through the multi-agent system.
        
        Args:
            user_message: The user's request
            
        Returns:
            dict: Contains 'answer', 'generated_code', and 'slide_context'
        """
        with trace_tool_call("multiagent_request", user_message=user_message[:100]):
                try:
                    add_trace_event("multiagent_start", user_message=user_message)
                    
                    # Get current slide context
                    slide_context = get_current_slide_context(force_refresh=True)
                    
                    # Enhanced message with slide context
                    enhanced_message = f"""CURRENT SLIDE CONTEXT:
    {slide_context}

    USER REQUEST: {user_message}

    INSTRUCTIONS:
    1. Analyze the user request and current slide state
    2. Determine if visual analysis is needed (call Vision Agent with slide image)
    3. Execute any required PowerPoint modifications through the Writing Agent
    4. Provide a comprehensive response to the user

    The Vision Agent can analyze slide visuals, and the Writing Agent can execute all PowerPoint operations.
    """
                    
                    add_trace_event("manager_execution", has_images=False)
                    
                    # Run the manager agent (no images by default - manager must explicitly get image if needed)
                    answer = self.manager_agent.run(enhanced_message)
                    
                    # Get updated slide context after operations
                    updated_context = get_current_slide_context(force_refresh=True)
                    
                    add_trace_event("multiagent_completed", success=True)
                    
                    return {
                        'answer': answer,
                        'generated_code': "# Multi-agent system executed successfully\n# Operations handled by specialized agents",
                        'slide_context': updated_context
                    }
                    
                except Exception as e:
                    add_trace_event("multiagent_error", error=str(e))
                    return {
                        'answer': f"Error in multi-agent system: {str(e)}",
                        'generated_code': f"# Error: {str(e)}",
                        'slide_context': "Error reading slide context"
                    }

# Global multi-agent system instance
multiagent_system = None

def get_multiagent_system():
    """Get or create the global multi-agent system instance."""
    global multiagent_system
    if multiagent_system is None:
        multiagent_system = MultiAgentPPTSystem()
    return multiagent_system

def run_multiagent_request(message: str) -> dict:
    """
    Run a request through the multi-agent PowerPoint system.
    
    Args:
        message: The user's message/request
        
    Returns:
        dict: Contains 'answer', 'generated_code', and 'slide_context'
    """
    system = get_multiagent_system()
    return system.process_request(message)

# Compatibility function for existing GUI
def run_agent_with_code_capture(message: str, images=None) -> dict:
    """
    Compatibility function that routes requests to the multi-agent system.
    This maintains compatibility with the existing GUI while using the new architecture.
    
    Args:
        message: The user's message/request
        images: Legacy parameter (images are now handled automatically)
        
    Returns:
        dict: Contains 'answer', 'generated_code', and 'slide_context'
    """
    # Note: images parameter is ignored as the multi-agent system handles vision automatically
    return run_multiagent_request(message)
