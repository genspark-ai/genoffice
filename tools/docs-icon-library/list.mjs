/**
 * The icons Docs offers under Insert > Icons, by category: the list a reviewer
 * reads. One icon per line, `name`, an optional `"Label"`, then `: keywords`.
 * The name is the Fluent UI System Icons file stem (`<name>_24_filled.svg`);
 * the label defaults to the name with its underscores as spaces. An icon sits
 * in one category; keywords are what else the search finds it by.
 *
 * Generic meanings only: no brand or product marks, no emblems of real
 * organizations. After editing, run `node tools/docs-icon-library/build.mjs`.
 */
export const ICON_LIST = {
  people: `
    person: user account profile
    people: users group two
    people_team "Team": group users staff
    people_community "Community": group society crowd
    person_add "Add person": invite new user
    person_circle "Person in circle": avatar profile account
    person_star "Star person": favorite favourite best employee
    person_heart "Person with heart": care love volunteer
    person_walking "Walking": pedestrian walk
    person_support "Support agent": help desk customer service headset
    person_voice "Speaking": talk voice speech
    person_lightbulb "Person with idea": idea creative thinker
    guest: visitor badge
    accessibility: accessible inclusion body figure
    diversity: inclusion variety equality
    emoji "Smile": happy face smiley
    emoji_laugh "Laugh": happy joy face
    emoji_sad "Sad": unhappy face
    emoji_angry "Angry": mad face
    emoji_surprise "Surprised": shock wow face
    hand_wave "Waving hand": hello hi goodbye greeting
    hand_right "Raised hand": stop palm volunteer
    hand_point "Pointing hand": finger click touch
    thumb_like "Thumbs up": like approve good yes
    thumb_dislike "Thumbs down": dislike disapprove bad no
  `,
  business: `
    briefcase: work job portfolio case
    handshake: deal agreement partnership
    organization "Organization chart": hierarchy structure org organisation
    contact_card "Contact card": business card id badge
    contact_card_group "Contact cards": address book directory
    target: goal aim objective bullseye
    target_arrow "Target with arrow": goal hit aim objective
    lightbulb "Light bulb": idea innovation tip
    rocket: launch startup growth
    gavel: law legal auction judge court
    scales "Scales of justice": law balance legal fairness
    vote: ballot election poll
    megaphone: announce marketing campaign
    presenter: presentation speaker talk
    clipboard_task "Task clipboard": checklist todo
    timeline: schedule roadmap history
    flowchart: process workflow diagram
    diagram: structure tree hierarchy
    gantt_chart "Gantt chart": project plan schedule
    thinking: thought idea consider
    brain: mind intelligence think
    puzzle_piece "Puzzle piece": solution fit part
  `,
  commerce: `
    cart "Shopping cart": trolley buy shop
    shopping_bag "Shopping bag": buy purchase store
    shopping_bag_percent "Discount bag": sale offer percent
    tag: label price
    tag_percent "Discount tag": sale offer percent
    money: cash banknote payment
    money_hand "Money in hand": pay cash donate
    money_calculator "Money and calculator": budget accounting finance
    coin_multiple "Coins": money change cash
    coin_stack "Coin stack": money savings cash
    currency_dollar_euro "Dollar and euro": exchange currency money
    wallet: money purse pay
    payment "Payment card": credit card debit pay
    receipt: bill invoice
    savings "Piggy bank": save money
    gift_card "Gift card": voucher coupon
    barcode_scanner "Barcode scanner": scan product code
    box: package parcel delivery
    ticket_diagonal "Ticket": pass coupon admission
    calculator: math compute sum
  `,
  communication: `
    mail: email envelope letter message
    mail_read "Opened mail": email read envelope
    mail_unread "Unread mail": email new envelope
    mail_inbox "Inbox": email tray
    send: paper plane submit message
    chat: message bubble speech talk
    chat_multiple "Chats": conversation messages bubbles
    chat_bubbles_question "Question and answer": faq help conversation
    comment: note remark bubble
    call "Phone call": telephone handset ring
    phone "Mobile phone": smartphone cell device
    voicemail: message recording
    mention "At sign": email address at
    megaphone_loud "Loud megaphone": announce shout broadcast
    alert "Bell": notification alarm
    alert_on "Ringing bell": notification alarm ring
    share: send export forward
    link: url chain hyperlink
    attach "Paper clip": attachment clip
    translate: language translation
    people_chat "Group chat": conversation discussion team
    news: newspaper article press
    walkie_talkie "Walkie-talkie": radio two-way
    mic "Microphone": record voice audio speak
  `,
  technology: `
    laptop: computer notebook pc
    desktop "Desktop computer": monitor screen pc
    desktop_tower "Tower computer": pc workstation
    tablet: device slate
    smartwatch: watch wearable
    tv "Television": screen display monitor
    keyboard: type keys input
    cursor "Pointer": mouse arrow
    server: rack host data center centre
    database: data storage cylinder
    cloud: storage online
    cloud_arrow_up "Cloud upload": upload backup
    cloud_arrow_down "Cloud download": download
    router: network modem internet
    wifi_1 "Wireless signal": wi-fi internet network
    cellular_data_1 "Signal bars": mobile network reception
    globe: world internet web earth
    code: programming developer brackets
    bug: insect error defect beetle
    bot "Robot": chatbot assistant automation ai
    brain_circuit "Artificial intelligence": ai machine learning
    plug_connected "Plug connected": power connect socket
    battery_10 "Full battery": power charge energy
    battery_charge "Battery charging": power energy
    hard_drive "Hard drive": disk storage
    print "Printer": printing paper
    headset_vr "VR headset": virtual reality goggles
    window "Window": app browser screen
    cube: 3d box block model
  `,
  analytics: `
    data_area "Area chart": graph statistics
    data_bar_horizontal "Bar chart": graph horizontal statistics
    data_bar_vertical "Column chart": graph bar vertical statistics
    data_bar_vertical_ascending "Rising columns": growth increase graph
    data_funnel "Funnel chart": conversion sales stages
    data_line "Line chart": graph trend
    data_pie "Pie chart": share proportion graph
    data_scatter "Scatter chart": plot points correlation
    data_trending "Trend chart": growth line graph
    data_waterfall "Waterfall chart": bridge graph
    chart_multiple "Charts": dashboard graphs report
    arrow_trending "Trending up": growth increase rise
    arrow_trending_down "Trending down": decline decrease fall
    poll: survey bars results
    gauge: meter dial speed performance
    top_speed "Speedometer": fast performance dial
    pulse: activity heartbeat signal wave
    table: grid spreadsheet rows columns
    filter: funnel refine
    math_formula "Formula": equation function fx
    autosum "Sum": sigma total add
    number_symbol "Number sign": hash pound hashtag
    text_percent "Percent": percentage rate
  `,
  office: `
    document: file page paper
    document_text "Text document": file page writing
    document_add "New document": file create
    document_edit "Edit document": file write pencil
    document_signature "Signed document": contract agreement
    folder: directory files
    folder_open "Open folder": directory files
    folder_zip "Zip folder": archive compressed
    archive "Archive box": storage records
    clipboard: paste board
    clipboard_task_list_ltr "Checklist": tasks todo clipboard
    note "Sticky note": memo post-it
    notebook: journal diary
    save "Floppy disk": save disk
    copy: duplicate
    cut "Scissors": cut trim
    pin: pushpin attach
    bookmark: mark favorite favourite ribbon
    pen: write ink
    edit "Pencil": write draw edit
    signature: sign autograph
    highlight "Highlighter": marker pen
    eraser: rubber delete
    calendar_ltr "Calendar": date schedule
    calendar_month "Month calendar": date schedule
    calendar_clock "Calendar and clock": appointment schedule time
    calendar_checkmark "Calendar checked": event confirmed date
    clock: time hour watch
    clock_alarm "Alarm clock": time wake reminder
    timer "Stopwatch": time countdown
    hourglass: time wait sand
    drawer: cabinet storage filing
  `,
  education: `
    hat_graduation "Graduation cap": mortarboard degree school university
    book: read closed
    book_open "Open book": read study
    book_open_lightbulb "Book with idea": learn knowledge study
    book_letter "Language book": dictionary alphabet reading
    book_globe "Geography book": world atlas
    library: books shelf
    backpack: school bag rucksack
    class: classroom lesson students
    teaching: teacher lesson board
    whiteboard: board draw classroom
    ruler: measure length
    beaker: chemistry lab science flask
    microscope: science lab biology
    molecule: chemistry atom science
    math_symbols "Math symbols": maths arithmetic plus minus
    pi: math maths constant
    quiz_new "Quiz": test exam question
    certificate: diploma award qualification
    earth: planet world globe geography
    planet: space saturn astronomy
    text_font "Letter A": alphabet font type
  `,
  medical: `
    stethoscope: doctor health examination
    patient: hospital bed health
    pill: medicine drug tablet capsule
    syringe: injection vaccine needle
    briefcase_medical "First aid kit": medical bag emergency
    heart_pulse "Heartbeat": heart rate cardiology health
    dentist "Tooth": dental teeth
    xray "X-ray": radiology bones scan
    temperature "Thermometer": fever heat
    sanitize "Hand sanitizer": hygiene clean wash sanitiser
    document_heart_pulse "Medical record": health chart patient file
    clipboard_pulse "Medical chart": health record
    document_pill "Prescription": medicine pharmacy
    eye: vision sight look
    glasses: spectacles vision optician
    sleep: rest night zzz
    plant_ragweed "Pollen": allergy hay fever
  `,
  transport: `
    vehicle_car "Car": automobile vehicle drive
    vehicle_cab "Taxi": cab car
    vehicle_bus "Bus": coach public transport
    vehicle_truck "Truck": lorry delivery
    vehicle_truck_profile "Delivery truck": van shipping lorry
    vehicle_bicycle "Bicycle": bike cycle cycling
    vehicle_motorcycle "Motorcycle": motorbike scooter
    vehicle_ship "Ship": boat ferry cruise sea
    vehicle_subway "Train": subway underground metro railway rail tube
    vehicle_tractor "Tractor": farm agriculture
    vehicle_car_parking "Parking": car park lot
    airplane "Airplane": aeroplane plane flight travel
    airplane_take_off "Take-off": departure flight plane
    airplane_landing "Landing": arrival flight plane
    gas_pump "Gas pump": fuel petrol station
    road: street route way
    road_cone "Traffic cone": roadworks construction caution
    street_sign "Street sign": signpost direction
    directions "Directions": route turn navigation
    map: location route plan
    location "Location pin": place marker map
    compass_northwest "Compass": direction navigation north
    luggage: suitcase baggage travel
  `,
  buildings: `
    home: house residence
    building: office block
    building_multiple "Buildings": offices blocks
    building_bank "Bank": finance columns institution
    building_factory "Factory": industry plant manufacturing
    building_government "Government building": city hall parliament
    building_retail "Store": shop retail market
    building_shop "Shop": store kiosk stall
    building_skyscraper "Skyscraper": tower high-rise
    building_townhouse "Townhouse": house terrace
    building_lighthouse "Lighthouse": coast beacon
    city: skyline town urban
    real_estate "Property for sale": house estate sign
    door: entrance exit
    conference_room "Meeting room": conference table
    bench: park seat
  `,
  nature: `
    leaf_one "Leaf": plant green eco
    leaf_two "Two leaves": plant sprout eco
    tree_deciduous "Tree": forest wood park
    plant_grass "Grass": lawn field
    clover: luck shamrock four-leaf
    earth_leaf "Green earth": environment eco sustainability
    water: waves sea ocean
    drop "Droplet": water rain liquid
    fire: flame hot burn
    mountain_trail "Mountains": hiking peak trail
    animal_cat "Cat": pet kitten
    animal_dog "Dog": pet puppy
    animal_rabbit "Rabbit": bunny pet
    animal_turtle "Turtle": tortoise slow
    animal_paw_print "Paw print": pet animal track
    weather_sunny "Sun": sunny clear day
    weather_moon "Moon": night crescent
    weather_cloudy "Cloudy": clouds overcast
    weather_partly_cloudy_day "Sun and cloud": partly cloudy
    weather_rain "Rain": rainy shower
    weather_drizzle "Drizzle": light rain
    weather_thunderstorm "Thunderstorm": lightning storm
    weather_snow "Snow": snowy cloud winter
    weather_snowflake "Snowflake": cold winter ice
    weather_fog "Fog": mist
    weather_squalls "Wind": squall gust breeze
    umbrella: rain cover protection
  `,
  food: `
    food "Knife and fork": cutlery restaurant meal eat
    food_apple "Apple": fruit healthy
    food_cake "Cake": birthday dessert slice
    food_carrot "Carrot": vegetable healthy
    food_chicken_leg "Chicken leg": meat drumstick
    food_egg "Egg": breakfast
    food_fish "Fish": seafood
    food_grains "Grain": wheat cereal bread
    food_pizza "Pizza": slice fast food
    bowl_chopsticks "Bowl and chopsticks": noodles rice
    bowl_salad "Salad bowl": vegetables healthy
    drink_beer "Beer": pint glass pub
    drink_coffee "Coffee": cup tea hot drink
    drink_margarita "Cocktail": margarita glass drink
    drink_to_go "Drink to go": takeaway cup straw soda
    drink_wine "Wine": glass drink
    spatula_spoon "Spatula and spoon": cooking utensils kitchen
    service_bell "Service bell": reception hotel restaurant
  `,
  sports: `
    sport "Sports": games tennis ball
    sport_american_football "American football": ball
    sport_baseball "Baseball": ball
    sport_basketball "Basketball": ball hoop
    sport_hockey "Hockey": stick puck
    sport_soccer "Soccer ball": football
    dumbbell: gym weights fitness exercise
    run "Running shoe": run runner jog exercise sneaker trainer
    steps "Stairs": steps staircase climb
    games "Game controller": gamepad play video game
    chess "Chess": pawn strategy game
    teddy "Teddy bear": toy child
    guitar: music instrument
    sport_cricket_ball "Cricket ball": ball
    tent: camping camp outdoors
    beach: umbrella seaside holiday
    swimming_pool "Swimming pool": swim water ladder
  `,
  media: `
    camera: photo photography
    image "Picture": photo landscape
    image_multiple "Pictures": photos gallery
    video "Video camera": film record
    filmstrip "Film strip": movie cinema
    movies_and_tv "Clapperboard": film movie cinema
    music_note_2 "Music notes": song audio melody
    headphones: audio listen music
    speaker_2 "Loudspeaker": sound volume audio
    speaker_mute "Muted speaker": sound off silent
    play: start triangle
    pause: hold
    stop: square end
    record: circle
    fast_forward "Fast forward": skip ahead
    rewind: back
    next "Next track": skip forward
    previous "Previous track": skip back
    paint_brush "Paintbrush": art paint
    paint_bucket "Paint bucket": fill color colour
    color "Palette": color colour paint art
    eyedropper: pipette color colour picker
    calligraphy_pen "Fountain pen": nib writing
    wand "Magic wand": effects auto
    closed_caption "Closed captions": subtitles cc
  `,
  arrows: `
    arrow_up "Arrow up": north
    arrow_down "Arrow down": south
    arrow_left "Arrow left": west back
    arrow_right "Arrow right": east forward next
    arrow_up_left "Arrow up left": northwest diagonal
    arrow_up_right "Arrow up right": northeast diagonal
    arrow_down_left "Arrow down left": southwest diagonal
    arrow_down_right "Arrow down right": southeast diagonal
    arrow_bidirectional_left_right "Arrow left and right": horizontal both ways
    arrow_bidirectional_up_down "Arrow up and down": vertical both ways
    arrow_clockwise "Clockwise arrow": refresh reload rotate
    arrow_counterclockwise "Counterclockwise arrow": anticlockwise rotate back
    arrow_sync "Sync arrows": refresh cycle update
    arrow_repeat_all "Repeat arrows": loop cycle
    arrow_shuffle "Shuffle arrows": random mix cross
    arrow_swap "Swap arrows": exchange switch
    arrow_undo "Undo arrow": back curve
    arrow_redo "Redo arrow": forward curve
    arrow_reply "Reply arrow": back respond
    arrow_forward "Forward arrow": share next
    arrow_turn_right "Turn right": bend direction
    arrow_exit "Exit arrow": leave out
    arrow_download "Download arrow": save down
    arrow_upload "Upload arrow": send up
    arrow_expand "Expand arrows": enlarge grow
    arrow_minimize "Shrink arrows": collapse reduce
    arrow_maximize "Maximize arrows": full screen enlarge
    arrow_move "Move arrows": drag four directions
    chevron_up "Chevron up": caret collapse
    chevron_down "Chevron down": caret expand
    chevron_left "Chevron left": back previous
    chevron_right "Chevron right": next forward
    caret_up "Triangle up": increase
    caret_down "Triangle down": decrease dropdown
    caret_left "Triangle left": back
    caret_right "Triangle right": next play
  `,
  signs: `
    checkmark "Check mark": tick done yes correct
    checkmark_circle "Check mark in circle": tick done success
    dismiss "Cross": close x no wrong
    dismiss_circle "Cross in circle": close error no
    add "Plus": add new positive
    add_circle "Plus in circle": add new
    subtract "Minus": remove negative dash
    subtract_circle "Minus in circle": remove
    question "Question mark": help ask
    question_circle "Question in circle": help faq
    info "Information": about details i
    warning: caution alert triangle exclamation
    error_circle "Error": exclamation alert problem
    important "Exclamation mark": attention priority
    prohibited: forbidden ban no entry
    heart: love like favorite favourite
    star: favorite favourite rating
    star_half "Half star": rating
    flag: marker report country
    circle: round dot shape
    square: box shape
    triangle: shape delta
    diamond: shape rhombus gem
    shapes: geometry circle square triangle
    flash "Lightning bolt": power energy fast electric
    number_circle_1 "Number 1": one first step
    number_circle_2 "Number 2": two second step
    number_circle_3 "Number 3": three third step
    bin_recycle "Recycling bin": recycle waste environment
    text_quote "Quotation marks": quote citation
    checkbox_checked "Checked box": tick checkbox done option
    crown: king queen royal premium
  `,
  interface: `
    search "Magnifying glass": find look zoom
    settings "Gear": cog options preferences
    options "Sliders": settings adjust controls
    more_horizontal "More": ellipsis dots menu
    navigation "Menu": hamburger lines
    grid: tiles layout squares
    list: lines items
    apps: tiles applications
    zoom_in "Zoom in": magnify enlarge plus
    zoom_out "Zoom out": reduce minus
    full_screen_maximize "Full screen": expand enlarge
    open "Open in new": external launch
    delete "Trash can": bin delete remove waste
    eye_off "Hidden": invisible hide
    toggle_right "Switch on": toggle
    panel_left "Left panel": sidebar layout
    sign_out "Sign out": log out exit leave
    history: recent clock past
    dark_theme "Dark theme": contrast night mode
    brightness_high "Brightness": light sun display
    textbox "Text box": field input
    layer: stack levels
  `,
  celebration: `
    gift: present box surprise
    gift_open "Opened gift": present unwrap
    balloon: party birthday
    trophy: cup winner award prize
    medal: award winner prize
    ribbon "Award ribbon": prize rosette
    ribbon_star "Star ribbon": award prize rosette
    star_emphasis "Shining star": sparkle highlight
    sparkle: shine magic new stars
    reward: prize award badge
    bow_tie "Bow": ribbon gift party
    checkmark_starburst "Seal of approval": verified badge quality
    calendar_star "Special date": event anniversary holiday
    ticket_horizontal "Event ticket": admission concert pass
  `,
  home: `
    couch: sofa furniture living room
    bed: bedroom sleep hotel furniture
    fireplace: hearth fire warm
    oven: cooker kitchen appliance
    washer "Washing machine": laundry appliance
    dishwasher: kitchen appliance
    showerhead "Shower": bathroom water
    broom: sweep clean
    clothes_hanger "Clothes hanger": wardrobe coat
    flashlight: torch light
    home_heart "Home with heart": family love care
    door_tag "Door hanger": do not disturb hotel
    bin_full "Full trash can": bin rubbish waste garbage
    wrench: spanner tool repair fix
    wrench_screwdriver "Wrench and screwdriver": spanner tools repair maintenance
    toolbox: tools kit repair
  `,
  security: `
    lock_closed "Padlock": locked secure private
    lock_open "Open padlock": unlocked
    lock_shield "Padlock and shield": protection secure
    shield: protection security defence
    shield_checkmark "Shield with check mark": tick protected safe verified
    shield_lock "Shield with lock": protected secure
    shield_error "Shield warning": risk threat alert
    key: access password unlock
    key_multiple "Keys": access
    password: pin code asterisks
    fingerprint: biometric identity touch
    person_lock "Locked account": private user
    document_lock "Locked document": confidential protected
    vault: safe bank secure
    camera_dome "Security camera": cctv surveillance
    video_security "Surveillance camera": cctv monitor
  `,
}
