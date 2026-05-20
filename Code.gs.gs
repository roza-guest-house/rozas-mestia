/**
 * Roza Booking Engine backend for Google Apps Script.
 *
 * Recommended setup:
 * - Bind this script to the "Roza Booking Engine" spreadsheet, OR
 * - Set Script Property SPREADSHEET_ID if you keep it as a standalone script.
 *
 * Endpoints:
 *   GET  /exec                       -> standalone PMS web app
 *   GET  ?action=ping
 *   GET  ?action=getSettings
 *   GET  ?action=getHomepageBaseRateSummary
 *   GET  ?action=checkAvailability&checkIn=YYYY-MM-DD&checkOut=YYYY-MM-DD&roomType=PEAKMV&guests=2
 *   GET  ?action=searchAvailabilityProducts&check_in=YYYY-MM-DD&check_out=YYYY-MM-DD&guests=2&bed_setup=Best%20available
 *   POST { action: 'checkAvailability', ... }
 *   POST { action: 'createRequest', ... }
 *   POST { action: 'createWebsiteBooking', ... }
 */

const SHEET_NAMES = {
  ROOM_TYPES: 'Room_Types',
  ROOMS: 'Rooms',
  BASE_RATES: 'Base_Rates',
  RATES: 'Rates',
  COMMERCIAL_CONTROLS: 'Commercial_Controls',
  EVENT_FLAGS: 'Event_Flags',
  COMPETITOR_TRACKER: 'Competitor_Tracker',
  RECOMMENDATION_ACTION_LOG: 'Recommendation_Action_Log',
  OTA_UPDATE_WORKFLOW: 'OTA_Update_Workflow',
  BOOKINGS: 'Bookings',
  BLOCKED_DATES: 'Blocked_Dates',
  AVAILABILITY_CACHE: 'Availability_Cache',
  REQUESTS: 'Requests',
  SETTINGS: 'Settings',
  DASHBOARD: 'Dashboard',
  BOOKING_NIGHTS: 'Booking_Nights',
  DAILY_STATS: 'Daily_Stats',
  OTB_SNAPSHOTS: 'OTB_Snapshots',
  WEBSITE_FEEDBACK: 'Website_Feedback'
};

const ROOM_TYPE_NAME_TO_ID = {
  'Classic Room (Shared Bathroom)': 'CLASSIC',
  'Peak Mountain View Room': 'PEAKMV',
  'Peak Mountain View Room with Balcony': 'PEAKBAL',
  'Peak Cottage with Mountain View': 'COTTAGE'
};

const SELLABLE_ROOM_PRODUCTS = [
  {
    product_id: 'CLASSIC_2',
    room_type_id: 'CLASSIC',
    product_name: 'Classic Room',
    product_label: 'Classic Room',
    standard_label: 'Classic room',
    capacity_label: '2 guests',
    product_summary: 'Simple Classic guest-house room for up to two guests using the valid Classic room inventory for your search.',
    bed_summary: 'Double or twin preference matched where valid',
    image_src: 'Classic/classic double .png',
    image_alt: 'Classic Room at Roza\'s Guest House',
    highlights: ['Shared bathroom', 'Classic house stay', 'Up to 2 guests'],
    min_guests: 1,
    max_guests: 2,
    occupancy_bucket: 2,
    supported_bed_setups: 'Double|Twin',
    candidate_room_codes: 'N2|N3|N4|N9|N10',
    sort_order: 10
  },
  {
    product_id: 'PEAKMV_2',
    room_type_id: 'PEAKMV',
    product_name: 'Peak Mountain View',
    product_label: 'Peak Mountain View',
    standard_label: 'Mountain-view room',
    capacity_label: '2 guests',
    product_summary: 'Private Peak Mountain View room for up to two guests using the valid two-guest inventory for your search.',
    bed_summary: 'Double or twin preference matched where valid',
    image_src: 'Peack Mountain View with Balcony/IMG_9124.PNG',
    image_alt: 'Peak Mountain View room at Roza\'s Guest House',
    highlights: ['Ensuite', 'Mountain view', 'Up to 2 guests'],
    min_guests: 1,
    max_guests: 2,
    occupancy_bucket: 2,
    supported_bed_setups: 'Double|Twin',
    candidate_room_codes: '5|6|7|8',
    sort_order: 20
  },
  {
    product_id: 'PEAKMV_3',
    room_type_id: 'PEAKMV',
    product_name: 'Peak Mountain View',
    product_label: 'Peak Mountain View',
    standard_label: 'Mountain-view room',
    capacity_label: '3 guests',
    product_summary: 'Peak Mountain View room for three guests using the valid triple-capable inventory.',
    bed_summary: 'Triple room',
    image_src: 'Peack Mountain View/peak triple .png',
    image_alt: 'Peak Mountain View room for 3 guests at Roza\'s Guest House',
    highlights: ['Ensuite', 'Mountain view', 'Up to 3 guests'],
    min_guests: 3,
    max_guests: 3,
    occupancy_bucket: 3,
    required_bed_setup: 'Triple',
    supported_bed_setups: 'Triple',
    candidate_room_codes: '1',
    sort_order: 21
  },
  {
    product_id: 'PEAKBAL_2',
    room_type_id: 'PEAKBAL',
    product_name: 'Peak Mountain View with Balcony',
    product_label: 'Peak Mountain View with Balcony',
    standard_label: 'Balcony room',
    capacity_label: '2 guests',
    product_summary: 'Private balcony room for up to two guests using the valid balcony inventory for your search.',
    bed_summary: 'Double or twin preference matched where valid',
    image_src: 'Peack Mountain View/peak double 2.png',
    image_alt: 'Peak Mountain View with Balcony at Roza\'s Guest House',
    highlights: ['Private balcony', 'Ensuite', 'Up to 2 guests'],
    min_guests: 1,
    max_guests: 2,
    occupancy_bucket: 2,
    supported_bed_setups: 'Double|Twin',
    candidate_room_codes: 'E|S|G|L|K|Y|R',
    sort_order: 30
  },
  {
    product_id: 'PEAKBAL_3',
    room_type_id: 'PEAKBAL',
    product_name: 'Peak Mountain View with Balcony',
    product_label: 'Peak Mountain View with Balcony',
    standard_label: 'Balcony room',
    capacity_label: '3 guests',
    product_summary: 'Balcony room for three guests using the valid triple-capable balcony inventory.',
    bed_summary: 'Triple room',
    image_src: 'Peack Mountain View with Balcony/Premium Triple .png',
    image_alt: 'Peak Mountain View with Balcony room for 3 guests at Roza\'s Guest House',
    highlights: ['Private balcony', 'Ensuite', 'Up to 3 guests'],
    min_guests: 3,
    max_guests: 3,
    occupancy_bucket: 3,
    required_bed_setup: 'Triple',
    supported_bed_setups: 'Triple',
    candidate_room_codes: 'G',
    sort_order: 31
  },
  {
    product_id: 'COTTAGE_4',
    room_type_id: 'COTTAGE',
    product_name: 'Peak Cottage',
    product_label: 'Peak Cottage',
    standard_label: 'Private cottage',
    capacity_label: 'Up to 4 guests',
    product_summary: 'Private cottage with one double room and one twin room, best for families or small groups.',
    bed_summary: 'Includes one double room and one twin room',
    image_src: 'cottage refined/IMG_9193.PNG',
    image_alt: 'Peak Cottage at Roza\'s Guest House',
    highlights: ['Private cottage', 'Double + twin', 'Up to 4 guests'],
    min_guests: 1,
    max_guests: 4,
    occupancy_bucket: 4,
    supported_bed_setups: 'Double|Twin',
    candidate_room_codes: 'A|C',
    sort_order: 40
  }
];

// Adjust these if your commercial rules change.
const INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE = {
  CLASSIC: 2,
  PEAKMV: 2,
  PEAKBAL: 2,
  COTTAGE: 4
};

const RATE_LADDER_PROPERTY_KEY = 'ROZA_RATE_LADDER_V1';
const RATE_LADDER_LEAD_ROOM_TYPE_ID = 'CLASSIC';
const RATE_BOARD_SORT_ORDER = ['CLASSIC', 'PEAKMV', 'PEAKBAL', 'COTTAGE'];
const HOMEPAGE_ROOM_DISPLAY_NAMES_BY_TYPE = {
  CLASSIC: 'Classic Room - simple, warm, best value',
  PEAKMV: 'Mountain View Room - wake up to the Caucasus',
  PEAKBAL: 'Balcony Room - open skies from your room',
  COTTAGE: 'Peak Cottage - your private mountain hideaway'
};
const PMS_ADMIN_EMAILS_PROPERTY_KEY = 'PMS_ADMIN_EMAILS';
const PMS_ADMIN_ACCESS_KEY_PROPERTY_KEY = 'PMS_ADMIN_ACCESS_KEY';
const PMS_ALLOW_OWNER_FALLBACK_PROPERTY_KEY = 'PMS_ALLOW_OWNER_FALLBACK';
const PMS_STANDALONE_URL_PROPERTY_KEY = 'PMS_STANDALONE_URL';
const PMS_ACCESS_CACHE_PREFIX = 'roza_pms_access_v1';
const PMS_ACCESS_SESSION_TTL_SECONDS = 21600;
const PMS_ADMIN_PIN_PROPERTY_KEY = 'PMS_ADMIN_PIN';
const PMS_PIN_SESSION_PROPERTY_PREFIX = 'PMS_PIN_SESSION_V1_';
const PMS_PIN_SESSION_CACHE_PREFIX = 'roza_pms_pin_session_v1';
const PMS_PIN_SESSION_TTL_SECONDS = 10 * 60 * 60;
const PMS_PIN_SESSION_CACHE_TTL_SECONDS = 21600;
const PUBLIC_RATE_LIMIT_CACHE_PREFIX = 'roza_public_limit_v1';
const PUBLIC_FORM_MIN_AGE_MS = 3500;
const PUBLIC_FORM_MAX_AGE_MS = 86400000;
const PUBLIC_SEARCH_WINDOW_SECONDS = 60;
const PUBLIC_SEARCH_MAX_REQUESTS = 30;
const PUBLIC_REQUEST_WINDOW_SECONDS = 3600;
const PUBLIC_REQUEST_MAX_REQUESTS = 6;
const PUBLIC_BOOKING_WINDOW_SECONDS = 3600;
const PUBLIC_BOOKING_MAX_REQUESTS = 4;
const PUBLIC_DUPLICATE_WINDOW_SECONDS = 900;
const PUBLIC_MAX_STAY_NIGHTS = 30;
const PUBLIC_MAX_LOOKAHEAD_DAYS = 365;

const BOOKING_STATUS_CONFIRMED = 'Confirmed';
const BOOKING_STATUS_IN_HOUSE = 'In House';
const BOOKING_STATUS_CHECKED_OUT = 'Checked Out';
const BOOKING_STATUS_CANCELLED = 'Cancelled';
const BOOKING_STATUS_NO_SHOW = 'No Show';
const BOOKING_STATUSES_COUNTED = [
  BOOKING_STATUS_CONFIRMED,
  BOOKING_STATUS_IN_HOUSE,
  BOOKING_STATUS_CHECKED_OUT
];
const FRONT_DESK_ACTIVE_BOOKING_STATUSES = [
  BOOKING_STATUS_CONFIRMED,
  BOOKING_STATUS_IN_HOUSE
];
const REQUEST_STATUS_INITIAL = 'New';
const REQUEST_STATUS_QUOTED = 'Quoted';
const REQUEST_STATUS_AWAITING_GUEST = 'Awaiting Guest';
const REQUEST_STATUS_CONVERTED = 'Converted';
const REQUEST_STATUS_LOST = 'Lost';
const REQUEST_STATUS_CANCELLED = 'Cancelled';
const REQUEST_STATUS_PENDING_VALUES = [
  REQUEST_STATUS_INITIAL,
  REQUEST_STATUS_QUOTED,
  REQUEST_STATUS_AWAITING_GUEST
];
const REQUEST_SOURCE_WEBSITE = 'website';
const DEFAULT_CURRENCY = 'EUR';
const DEFAULT_REPORTING_CURRENCY = 'GBP';
const BOOKING_CURRENCY_VALUES = ['GBP', 'EUR', 'USD', 'GEL'];
const FX_CACHE_TTL_SECONDS = 21600;
const FX_FALLBACK_TO_GBP = {
  GBP: 1,
  EUR: 0.86,
  USD: 0.79,
  GEL: 0.29
};
const DEFAULT_PROPERTY_NAME = 'Roza\'s Guest House';
const DEFAULT_PROPERTY_ADDRESS = '9 Erekle Parjiani I Alley, Mestia 3200, Georgia';
const DEFAULT_PROPERTY_PHONE = '+995599641455';
const DEFAULT_PROPERTY_EMAIL = 'rozashukvani@gmail.com';
const WEBSITE_DEFAULT_CHECK_IN_TIME = '15:00';
const WEBSITE_DEFAULT_CHECK_OUT_TIME = '11:00';
const WEBSITE_EMAIL_FOLLOW_UP_HANDLER = 'processPendingWebsiteBookingEmails';
const WEBSITE_EMAIL_FOLLOW_UP_TRIGGER_DELAY_MS = 60 * 1000;
const WEBSITE_EMAIL_FOLLOW_UP_BATCH_LIMIT = 10;
const WEBSITE_EMAIL_FOLLOW_UP_SENDING_STALE_MINUTES = 20;
const BOOKING_ID_COUNTER_PROPERTY_PREFIX = 'ROZA_BOOKING_ID_COUNTER_BKG_';
const DASHBOARD_SEARCH_CELL = 'B3';
const OTB_SNAPSHOT_DAYS_AHEAD = 365;
const BOOKING_SOURCE_VALUES = [
  'Booking.com',
  'Airbnb',
  'Direct Website',
  'WhatsApp',
  'Phone',
  'Walk-in',
  'Manual'
];
const PAYMENT_STATUS_VALUES = [
  'Pay at Property',
  'Unpaid',
  'Partially Paid',
  'Paid',
  'Refunded'
];
const PAYMENT_METHOD_VALUES = [
  'Cash',
  'Bank Transfer',
  'Card',
  'Booking.com Virtual Card',
  'Airbnb Payout',
  'Online Payment',
  'Other'
];
const BED_SETUP_VALUES = [
  'Best available',
  'Double',
  'Twin',
  'Triple'
];
const COMMERCIAL_RULE_TYPE_VALUES = ['seasonal', 'special'];
const EVENT_FLAG_TYPE_VALUES = ['Holiday', 'Local Event', 'Festival', 'Peak Season', 'School Break', 'Market Compression', 'Manual Commercial Flag', 'Other'];
const EVENT_FLAG_IMPACT_VALUES = ['Low', 'Medium', 'High', 'Very High'];
const COMPETITOR_TRACKER_AVAILABILITY_VALUES = ['Available', 'Sold Out'];
const COMPETITOR_TRACKER_BREAKFAST_VALUES = ['Yes', 'No', 'Unknown'];
const COMPETITOR_TRACKER_SOURCE_VALUES = ['Booking.com', 'Airbnb', 'Direct website', 'Other'];
const RECOMMENDATION_ACTION_TAKEN_VALUES = [
  'Held price',
  'Raised ' + DEFAULT_CURRENCY + ' 5',
  'Raised ' + DEFAULT_CURRENCY + ' 10',
  'Raised ' + DEFAULT_CURRENCY + ' 20',
  'Dropped ' + DEFAULT_CURRENCY + ' 5',
  'Dropped ' + DEFAULT_CURRENCY + ' 10',
  'Raised custom',
  'Dropped custom',
  'Set exact price',
  'Changed manually',
  'No action',
  'Reviewed only'
];
const RECOMMENDATION_OTA_UPDATED_VALUES = ['Yes', 'No', 'Not relevant'];
const RECOMMENDATION_OVERBOOKING_CHANGE_VALUES = ['No change', 'Increased by 1', 'Reduced by 1', 'Other'];
const OTA_WORKFLOW_STATUS_VALUES = ['Pending', 'Updated', 'Not Relevant'];
const REQUESTS_HEADERS = [
  'request_id',
  'created_at',
  'guest_name',
  'guest_phone',
  'guest_email',
  'check_in',
  'check_out',
  'room_type_id',
  'room_type_name',
  'bed_setup',
  'guests',
  'estimated_price',
  'request_status',
  'request_source',
  'ack_email_status',
  'ack_email_sent_at',
  'ack_email_error',
  'notes',
  'booking_id',
  'converted_to_booking_at',
  'conversion_status',
  'assigned_source',
  'assigned_booking_value'
];
const BOOKINGS_HEADERS = [
  'booking_id',
  'created_at',
  'request_id',
  'source',
  'source_detail',
  'guest_name',
  'guest_phone',
  'guest_email',
  'country',
  'check_in',
  'check_in_time',
  'check_out',
  'check_out_time',
  'actual_check_out_date',
  'nights',
  'room_type_id',
  'room_type_name',
  'room_identifier',
  'bed_setup',
  'adults',
  'children',
  'guests',
  'qty_rooms',
  'status',
  'booking_value',
  'booking_value_original',
  'booking_currency',
  'fx_rate_to_gbp',
  'booking_value_gbp',
  'pricing_source',
  'pricing_reference_id',
  'amount_paid',
  'payment_method',
  'tax_amount',
  'payment_status',
  'payment_notes',
  'balance_due',
  'currency',
  'cancelled_at',
  'cancel_reason',
  'converted_from_request_at',
  'guest_preferences',
  'notes',
  'internal_notes'
];
const BLOCKED_DATES_HEADERS = [
  'blocked_id',
  'date',
  'room_type_id',
  'qty_blocked',
  'reason',
  'status',
  'notes'
];
const RATES_HEADERS = [
  'date',
  'room_type_id',
  'base_rate',
  'extra_guest_fee',
  'min_stay',
  'status',
  'notes'
];
const BASE_RATES_HEADERS = [
  'room_type_id',
  'room_type_name',
  'base_rate',
  'extra_guest_fee',
  'public_reference_price',
  'direct_discount_type',
  'direct_discount_value',
  'active',
  'updated_at'
];
const COMMERCIAL_CONTROLS_HEADERS = [
  'control_id',
  'room_type_id',
  'room_type_name',
  'rule_type',
  'start_date',
  'end_date',
  'override_price',
  'public_reference_price',
  'direct_discount_type',
  'direct_discount_value',
  'overbooking_allowance',
  'active',
  'note',
  'created_at',
  'updated_at'
];
const EVENT_FLAGS_HEADERS = [
  'event_id',
  'event_name',
  'start_date',
  'end_date',
  'event_type',
  'impact_level',
  'note',
  'active',
  'updated_at'
];
const COMPETITOR_TRACKER_HEADERS = [
  'entry_id',
  'created_at',
  'date',
  'competitor_name',
  'matched_roza_room_type_id',
  'matched_roza_room_type',
  'competitor_room_description',
  'competitor_price',
  'available_or_sold_out',
  'breakfast_included',
  'source_checked',
  'note'
];
const RECOMMENDATION_ACTION_LOG_HEADERS = [
  'log_id',
  'date',
  'room_type_id',
  'room_type_name',
  'recommendation_action',
  'recommendation_confidence',
  'recommendation_reason_summary',
  'roza_old_price',
  'action_taken',
  'roza_new_price',
  'overbooking_change',
  'ota_updated',
  'operator_note',
  'created_at'
];
const OTA_UPDATE_WORKFLOW_HEADERS = [
  'ota_update_id',
  'date',
  'room_type_id',
  'room_type_name',
  'booking_com_status',
  'airbnb_status',
  'linked_recommendation_action',
  'linked_action_taken',
  'roza_price_after_change',
  'note',
  'updated_at'
];
const BOOKING_NIGHTS_HEADERS = [
  'stay_date',
  'stay_month',
  'booking_id',
  'request_id',
  'booking_created_at',
  'source',
  'source_detail',
  'booking_status',
  'guest_name',
  'room_type_id',
  'room_type_name',
  'qty_rooms',
  'room_nights',
  'nightly_room_revenue',
  'booking_value',
  'check_in',
  'check_out',
  'nights',
  'currency'
];
const DAILY_STATS_HEADERS = [
  'stay_date',
  'stay_month',
  'year',
  'month',
  'room_type_id',
  'room_type_name',
  'inventory_total',
  'blocked_rooms',
  'rooms_available_to_sell',
  'rooms_sold',
  'room_nights',
  'bookings_count',
  'room_revenue',
  'occupancy_pct',
  'adr',
  'revpar'
];
const ROOMS_HEADERS = [
  'room_id',
  'room_code',
  'room_name',
  'room_type_id',
  'room_type_name',
  'default_setup',
  'allowed_setups',
  'active',
  'sort_order',
  'notes'
];
const OTB_SNAPSHOTS_HEADERS = [
  'snapshot_date',
  'stay_date',
  'stay_month',
  'room_type_id',
  'room_type_name',
  'rooms_available_to_sell',
  'rooms_sold_otb',
  'room_revenue_otb',
  'occupancy_pct_otb',
  'adr_otb',
  'revpar_otb'
];
const WEBSITE_FEEDBACK_HEADERS = [
  'created_at',
  'booking_id',
  'guest_name',
  'guest_email',
  'website_score',
  'booking_process_score',
  'feedback_note',
  'client_session_id',
  'page_url',
  'user_agent'
];
const WEBSITE_FEEDBACK_DASHBOARD_HEADERS = [
  'created_at',
  'booking_id',
  'guest_name',
  'guest_email',
  'website_score',
  'booking_process_score',
  'feedback_note'
];
const WEBSITE_FEEDBACK_RECENT_BOOKING_WINDOW_MS = 24 * 60 * 60 * 1000;
const DEFAULT_ROOM_TYPES_SEED = [
  {
    room_type_id: 'PEAKBAL',
    room_type_name: 'Peak Mountain View Room with Balcony',
    inventory_total: 7,
    max_guests: 3,
    active: 'Yes'
  },
  {
    room_type_id: 'PEAKMV',
    room_type_name: 'Peak Mountain View Room',
    inventory_total: 5,
    max_guests: 3,
    active: 'Yes'
  },
  {
    room_type_id: 'CLASSIC',
    room_type_name: 'Classic Room (Shared Bathroom)',
    inventory_total: 5,
    max_guests: 2,
    active: 'Yes'
  },
  {
    room_type_id: 'COTTAGE',
    room_type_name: 'Peak Cottage with Mountain View',
    inventory_total: 2,
    max_guests: 4,
    active: 'Yes'
  }
];
const DEFAULT_ROOMS_MASTER = [
  { room_id: 'PEAKBAL-E', room_code: 'E', room_name: 'Peak Balcony E', room_type_id: 'PEAKBAL', room_type_name: 'Peak Mountain View Room with Balcony', default_setup: 'Double', allowed_setups: 'Double', active: 'Yes', sort_order: 1, notes: '' },
  { room_id: 'PEAKBAL-S', room_code: 'S', room_name: 'Peak Balcony S', room_type_id: 'PEAKBAL', room_type_name: 'Peak Mountain View Room with Balcony', default_setup: 'Double', allowed_setups: 'Double', active: 'Yes', sort_order: 2, notes: '' },
  { room_id: 'PEAKBAL-G', room_code: 'G', room_name: 'Peak Balcony G', room_type_id: 'PEAKBAL', room_type_name: 'Peak Mountain View Room with Balcony', default_setup: 'Triple', allowed_setups: 'Triple|Twin|Double', active: 'Yes', sort_order: 3, notes: '' },
  { room_id: 'PEAKBAL-L', room_code: 'L', room_name: 'Peak Balcony L', room_type_id: 'PEAKBAL', room_type_name: 'Peak Mountain View Room with Balcony', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 4, notes: '' },
  { room_id: 'PEAKBAL-Y', room_code: 'Y', room_name: 'Peak Balcony Y', room_type_id: 'PEAKBAL', room_type_name: 'Peak Mountain View Room with Balcony', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 5, notes: '' },
  { room_id: 'PEAKBAL-K', room_code: 'K', room_name: 'Peak Balcony K', room_type_id: 'PEAKBAL', room_type_name: 'Peak Mountain View Room with Balcony', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 6, notes: '' },
  { room_id: 'PEAKBAL-R', room_code: 'R', room_name: 'Peak Balcony R', room_type_id: 'PEAKBAL', room_type_name: 'Peak Mountain View Room with Balcony', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 7, notes: '' },
  { room_id: 'PEAK-N1', room_code: '1', room_name: 'Peak 1', room_type_id: 'PEAKMV', room_type_name: 'Peak Mountain View Room', default_setup: 'Triple', allowed_setups: 'Triple|Twin|Double', active: 'Yes', sort_order: 8, notes: '' },
  { room_id: 'PEAK-N5', room_code: '5', room_name: 'Peak 5', room_type_id: 'PEAKMV', room_type_name: 'Peak Mountain View Room', default_setup: 'Double', allowed_setups: 'Double', active: 'Yes', sort_order: 9, notes: '' },
  { room_id: 'PEAK-N6', room_code: '6', room_name: 'Peak 6', room_type_id: 'PEAKMV', room_type_name: 'Peak Mountain View Room', default_setup: 'Single / Queen', allowed_setups: 'Single / Queen', active: 'Yes', sort_order: 10, notes: '' },
  { room_id: 'PEAK-N7', room_code: '7', room_name: 'Peak 7', room_type_id: 'PEAKMV', room_type_name: 'Peak Mountain View Room', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 11, notes: '' },
  { room_id: 'PEAK-N8', room_code: '8', room_name: 'Peak 8', room_type_id: 'PEAKMV', room_type_name: 'Peak Mountain View Room', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 12, notes: '' },
  { room_id: 'CLASSIC-N2', room_code: 'N2', room_name: 'Classic N2', room_type_id: 'CLASSIC', room_type_name: 'Classic Room (Shared Bathroom)', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 13, notes: 'Permanent Twin' },
  { room_id: 'CLASSIC-N3', room_code: 'N3', room_name: 'Classic N3', room_type_id: 'CLASSIC', room_type_name: 'Classic Room (Shared Bathroom)', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 14, notes: 'Permanent Twin' },
  { room_id: 'CLASSIC-N4', room_code: 'N4', room_name: 'Classic N4', room_type_id: 'CLASSIC', room_type_name: 'Classic Room (Shared Bathroom)', default_setup: 'Twin', allowed_setups: 'Twin', active: 'Yes', sort_order: 15, notes: 'Permanent Twin' },
  { room_id: 'CLASSIC-N9', room_code: 'N9', room_name: 'Classic N9', room_type_id: 'CLASSIC', room_type_name: 'Classic Room (Shared Bathroom)', default_setup: 'Double', allowed_setups: 'Double', active: 'Yes', sort_order: 16, notes: '' },
  { room_id: 'CLASSIC-N10', room_code: 'N10', room_name: 'Classic N10', room_type_id: 'CLASSIC', room_type_name: 'Classic Room (Shared Bathroom)', default_setup: 'Double', allowed_setups: 'Double', active: 'Yes', sort_order: 17, notes: '' },
  { room_id: 'COTTAGE-A', room_code: 'A', room_name: 'Cottage A', room_type_id: 'COTTAGE', room_type_name: 'Peak Cottage with Mountain View', default_setup: '', allowed_setups: '', active: 'Yes', sort_order: 18, notes: '' },
  { room_id: 'COTTAGE-C', room_code: 'C', room_name: 'Cottage C', room_type_id: 'COTTAGE', room_type_name: 'Peak Cottage with Mountain View', default_setup: '', allowed_setups: '', active: 'Yes', sort_order: 19, notes: '' }
];

const ROOM_MASTER_COMPATIBILITY_BY_TYPE_AND_CODE = {
  PEAKBAL: {
    E: { roomCode: 'E', roomName: 'Peak Balcony E', defaultSetup: 'Double', allowedSetups: 'Double' },
    S: { roomCode: 'S', roomName: 'Peak Balcony S', defaultSetup: 'Double', allowedSetups: 'Double' },
    G: { roomCode: 'G', roomName: 'Peak Balcony G', defaultSetup: 'Triple', allowedSetups: 'Triple|Twin|Double' },
    L: { roomCode: 'L', roomName: 'Peak Balcony L', defaultSetup: 'Twin', allowedSetups: 'Twin' },
    Y: { roomCode: 'Y', roomName: 'Peak Balcony Y', defaultSetup: 'Twin', allowedSetups: 'Twin' },
    K: { roomCode: 'K', roomName: 'Peak Balcony K', defaultSetup: 'Twin', allowedSetups: 'Twin' },
    R: { roomCode: 'R', roomName: 'Peak Balcony R', defaultSetup: 'Twin', allowedSetups: 'Twin' }
  },
  PEAKMV: {
    '1': { roomCode: '1', roomName: 'Peak 1', defaultSetup: 'Triple', allowedSetups: 'Triple|Twin|Double' },
    '5': { roomCode: '5', roomName: 'Peak 5', defaultSetup: 'Double', allowedSetups: 'Double' },
    '6': { roomCode: '6', roomName: 'Peak 6', defaultSetup: 'Single / Queen', allowedSetups: 'Single / Queen' },
    '7': { roomCode: '7', roomName: 'Peak 7', defaultSetup: 'Twin', allowedSetups: 'Twin' },
    '8': { roomCode: '8', roomName: 'Peak 8', defaultSetup: 'Twin', allowedSetups: 'Twin' }
  }
};

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = String(params.action || '').trim();
  const requestedView = String(params.view || params.mode || params.app || '').trim().toLowerCase();
  const uiMode = requestedView === 'staff' ? 'staff' : 'web';
  const isPmsRequest = !action || requestedView === 'pms' || requestedView === 'staff';
  try {
    if (isPmsRequest) {
      let accessContext = null;
      try {
        accessContext = requirePmsAdminAccess_(params);
      } catch (accessError) {
        if (!isPmsAccessError_(accessError) || !hasPmsAdminPinConfigured_()) {
          throw accessError;
        }
        accessContext = {
          authorized: true,
          via: 'pin-login',
          activeEmail: '',
          accessKey: ''
        };
      }
      return renderPmsApp_(uiMode, accessContext);
    }
    return dispatchPublicWebAction_(action, params, 'GET');
  } catch (error) {
    if (isPmsRequest) {
      if (isPmsAccessError_(error)) {
        return renderPmsUnauthorizedPage_(error);
      }
      return renderPmsErrorPage_(error);
    }
    return jsonResponse_({ ok: false, error: error.message });
  }
}

function doPost(e) {
  try {
    const payload = parseIncomingPayload_(e);
    const action = String(payload.action || '').trim();

    if (!action) {
      throw new Error('Missing action in POST body.');
    }
    return dispatchPublicWebAction_(action, payload, 'POST');
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message });
  }
}

function dispatchPublicWebAction_(action, payload, method) {
  const normalizedAction = String(action || '').trim();
  const requestMethod = String(method || 'GET').trim().toUpperCase();
  const effectivePayload = payload || {};

  if (normalizedAction === 'ping') {
    return jsonResponse_({ ok: true, action: 'ping', message: 'Roza Booking Engine backend is live.' });
  }

  if (normalizedAction === 'getSettings') {
    return jsonResponse_({ ok: true, action: 'getSettings', data: getPublicWebsiteSettings_() });
  }

  if (normalizedAction === 'getHomepageBaseRateSummary') {
    try {
      return jsonResponse_({ ok: true, action: 'getHomepageBaseRateSummary', data: getHomepageBaseRateSummary_() });
    } catch (error) {
      return jsonResponse_({ ok: false, action: 'getHomepageBaseRateSummary', error: 'Homepage base rates are temporarily unavailable.' });
    }
  }

  if (normalizedAction === 'checkAvailability') {
    enforcePublicSearchProtection_(effectivePayload);
    return jsonResponse_({ ok: true, action: 'checkAvailability', data: checkAvailability(effectivePayload) });
  }

  if (normalizedAction === 'searchAvailabilityProducts') {
    enforcePublicSearchProtection_(effectivePayload);
    return jsonResponse_({ ok: true, action: 'searchAvailabilityProducts', data: searchAvailabilityProducts(effectivePayload) });
  }

  if (requestMethod === 'POST' && normalizedAction === 'createRequest') {
    enforcePublicSubmissionProtection_(effectivePayload, 'request');
    effectivePayload._publicSubmissionProtectionApplied = true;
    return jsonResponse_({ ok: true, action: 'createRequest', data: createRequest(effectivePayload) });
  }

  if (requestMethod === 'POST' && normalizedAction === 'createWebsiteBooking') {
    enforcePublicSubmissionProtection_(effectivePayload, 'booking');
    effectivePayload._publicSubmissionProtectionApplied = true;
    return jsonResponse_({ ok: true, action: 'createWebsiteBooking', data: createWebsiteBooking(effectivePayload) });
  }

  if (requestMethod === 'POST' && normalizedAction === 'sendWebsiteBookingEmailFollowUp') {
    enforcePublicEmailFollowUpProtection_(effectivePayload);
    return jsonResponse_({ ok: true, action: 'sendWebsiteBookingEmailFollowUp', data: sendWebsiteBookingEmailFollowUp(effectivePayload) });
  }

  if (requestMethod === 'POST' && normalizedAction === 'saveWebsiteBookingFeedback') {
    return jsonResponse_({ ok: true, action: 'saveWebsiteBookingFeedback', data: saveWebsiteBookingFeedback(effectivePayload) });
  }

  throw new Error('Unsupported ' + requestMethod + ' action: ' + normalizedAction);
}

function renderPmsErrorPage_(error) {
  const message = String(error && error.message || error || 'Unknown PMS startup error.').trim();
  const stack = String(error && error.stack || '').trim();
  const spreadsheetHint = message.indexOf('SPREADSHEET_ID') !== -1 || message.indexOf('spreadsheet') !== -1
    ? '<p style="margin:0;color:#667267;">Check the Apps Script <strong>SPREADSHEET_ID</strong> script property and confirm the deployment account can open the PMS spreadsheet.</p>'
    : '';
  const html = '' +
    '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Roza PMS Error</title>' +
    '<style>' +
      'body{margin:0;padding:32px;background:#f5f1e7;color:#1b241d;font:14px/1.6 Verdana,sans-serif;}' +
      '.panel{max-width:920px;margin:0 auto;background:#fff;border:1px solid rgba(27,52,40,.12);border-radius:22px;padding:28px;box-shadow:0 18px 48px rgba(23,36,29,.08);}' +
      'h1{margin:0 0 10px;font:600 30px/1.05 Georgia,serif;color:#183527;}' +
      'p{margin:0 0 12px;}' +
      'pre{margin:16px 0 0;padding:16px;border-radius:14px;background:#f6f3eb;white-space:pre-wrap;word-break:break-word;overflow:auto;font:12px/1.6 Consolas,monospace;color:#234a39;}' +
      '.eyebrow{text-transform:uppercase;letter-spacing:.12em;font-size:11px;color:#667267;margin-bottom:8px;}' +
    '</style></head><body><div class="panel">' +
      '<div class="eyebrow">Roza Guest House PMS</div>' +
      '<h1>PMS could not open</h1>' +
      '<p><strong>Error:</strong> ' + escapeHtmlForTemplate_(message) + '</p>' +
      spreadsheetHint +
      (stack ? '<pre>' + escapeHtmlForTemplate_(stack) + '</pre>' : '') +
    '</div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Roza PMS Error');
}

function renderPmsUnauthorizedPage_(error) {
  const configuredEmails = getPmsAdminAllowlist_();
  const hasAccessKey = !!getPmsAdminAccessKey_();
  const message = String(error && error.message || 'PMS access is restricted.').trim();
  const instructions = hasAccessKey
    ? 'Open the PMS using the protected access link, or ask the operator to approve your admin email in Apps Script.'
    : 'Set PMS_ADMIN_EMAILS or PMS_ADMIN_ACCESS_KEY in Apps Script script properties before exposing the PMS web app. Only use PMS_ALLOW_OWNER_FALLBACK=true for temporary owner-only testing.';
  const html = '' +
    '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<title>Roza PMS Access Required</title>' +
    '<style>' +
      'body{margin:0;padding:32px;background:#f5f1e7;color:#1b241d;font:14px/1.6 Verdana,sans-serif;}' +
      '.panel{max-width:760px;margin:0 auto;background:#fff;border:1px solid rgba(27,52,40,.12);border-radius:22px;padding:28px;box-shadow:0 18px 48px rgba(23,36,29,.08);}' +
      'h1{margin:0 0 10px;font:600 30px/1.05 Georgia,serif;color:#183527;}' +
      'p{margin:0 0 12px;}' +
      '.eyebrow{text-transform:uppercase;letter-spacing:.12em;font-size:11px;color:#667267;margin-bottom:8px;}' +
      '.meta{color:#667267;font-size:13px;}' +
    '</style></head><body><div class="panel">' +
      '<div class="eyebrow">Roza Guest House PMS</div>' +
      '<h1>Access required</h1>' +
      '<p><strong>' + escapeHtmlForTemplate_(message) + '</strong></p>' +
      '<p>' + escapeHtmlForTemplate_(instructions) + '</p>' +
      '<p class="meta">Allowlist configured: ' + escapeHtmlForTemplate_(configuredEmails.length ? 'Yes' : 'No') + ' | Access key configured: ' + escapeHtmlForTemplate_(hasAccessKey ? 'Yes' : 'No') + '</p>' +
    '</div></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('Roza PMS Access Required');
}

function getPmsAdminAllowlist_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PMS_ADMIN_EMAILS_PROPERTY_KEY) || '';
  return raw.split(/[\n,;]+/).map(function(email) {
    return String(email || '').trim().toLowerCase();
  }).filter(Boolean);
}

function getPmsAdminAccessKey_() {
  return String(PropertiesService.getScriptProperties().getProperty(PMS_ADMIN_ACCESS_KEY_PROPERTY_KEY) || '').trim();
}

function getPmsAllowOwnerFallback_() {
  const raw = String(PropertiesService.getScriptProperties().getProperty(PMS_ALLOW_OWNER_FALLBACK_PROPERTY_KEY) || '').trim().toLowerCase();
  return raw === '1' || raw === 'true' || raw === 'yes' || raw === 'on';
}

function getPmsTemporaryUserKey_() {
  try {
    return String(Session.getTemporaryActiveUserKey() || '').trim();
  } catch (error) {
    return '';
  }
}

function getPmsActiveUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  } catch (error) {
    return '';
  }
}

function getPmsEffectiveUserEmail_() {
  try {
    return String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  } catch (error) {
    return '';
  }
}

function extractPmsAccessKey_(input) {
  const payload = input || {};
  return String(
    payload._pmsAccessKey ||
    payload.pms_access_key ||
    payload.access_key ||
    payload.access ||
    payload.key ||
    ''
  ).trim();
}

function extractPmsSessionToken_(input) {
  const payload = input || {};
  return String(
    payload._pmsSessionToken ||
    payload.pmsSessionToken ||
    payload.session_token ||
    payload.sessionToken ||
    ''
  ).trim();
}

function getPmsAdminPin_() {
  const scriptPropertyPin = String(
    PropertiesService.getScriptProperties().getProperty(PMS_ADMIN_PIN_PROPERTY_KEY) || ''
  ).trim();
  if (scriptPropertyPin) return scriptPropertyPin;

  try {
    return String(getSettingValue_(getSettings(), [PMS_ADMIN_PIN_PROPERTY_KEY, 'pms_admin_pin'], '') || '').trim();
  } catch (error) {
    return '';
  }
}

function hasPmsAdminPinConfigured_() {
  return !!getPmsAdminPin_();
}

function buildPmsPinSessionCacheKey_(token) {
  return PMS_PIN_SESSION_CACHE_PREFIX + ':' + hashIdentifier_(token);
}

function buildPmsPinSessionPropertyKey_(token) {
  return PMS_PIN_SESSION_PROPERTY_PREFIX + hashIdentifier_(token);
}

function createPmsPinSessionToken_() {
  return Utilities.getUuid() + '.' + Utilities.getUuid();
}

function writePmsPinSession_(token) {
  const sessionToken = String(token || '').trim();
  if (!sessionToken) throw new Error('PMS session token could not be created.');
  const now = new Date();
  const expiresAt = new Date(now.getTime() + PMS_PIN_SESSION_TTL_SECONDS * 1000);
  const session = {
    valid: true,
    via: 'pin',
    activeEmail: getPmsActiveUserEmail_(),
    effectiveEmail: getPmsEffectiveUserEmail_(),
    createdAt: now.toISOString(),
    expiresAt: expiresAt.toISOString(),
    expiresAtMs: expiresAt.getTime()
  };
  const encoded = JSON.stringify(session);
  PropertiesService.getScriptProperties().setProperty(buildPmsPinSessionPropertyKey_(sessionToken), encoded);
  try {
    CacheService.getScriptCache().put(
      buildPmsPinSessionCacheKey_(sessionToken),
      encoded,
      Math.min(PMS_PIN_SESSION_TTL_SECONDS, PMS_PIN_SESSION_CACHE_TTL_SECONDS)
    );
  } catch (error) {
    // Script Properties remains the source of truth for the 10-hour expiry.
  }
  return session;
}

function deletePmsPinSession_(token) {
  const sessionToken = String(token || '').trim();
  if (!sessionToken) return;
  const propertyKey = buildPmsPinSessionPropertyKey_(sessionToken);
  const cacheKey = buildPmsPinSessionCacheKey_(sessionToken);
  try {
    CacheService.getScriptCache().remove(cacheKey);
  } catch (error) {
    // Logout should still clear the persistent token record.
  }
  PropertiesService.getScriptProperties().deleteProperty(propertyKey);
}

function parsePmsPinSession_(raw, token) {
  if (!raw) return null;
  try {
    const session = JSON.parse(raw);
    const expiresAtMs = Number(session && session.expiresAtMs || 0);
    if (!session || !session.valid || !expiresAtMs || expiresAtMs <= Date.now()) {
      deletePmsPinSession_(token);
      return null;
    }
    return session;
  } catch (error) {
    deletePmsPinSession_(token);
    return null;
  }
}

function readPmsPinSession_(token) {
  const sessionToken = String(token || '').trim();
  if (!sessionToken) return null;
  try {
    const cached = CacheService.getScriptCache().get(buildPmsPinSessionCacheKey_(sessionToken));
    const cachedSession = parsePmsPinSession_(cached, sessionToken);
    if (cachedSession) return cachedSession;
  } catch (error) {
    // Fall back to Script Properties below.
  }

  const raw = PropertiesService.getScriptProperties().getProperty(buildPmsPinSessionPropertyKey_(sessionToken));
  const session = parsePmsPinSession_(raw, sessionToken);
  if (session) {
    try {
      CacheService.getScriptCache().put(
        buildPmsPinSessionCacheKey_(sessionToken),
        JSON.stringify(session),
        Math.min(PMS_PIN_SESSION_TTL_SECONDS, PMS_PIN_SESSION_CACHE_TTL_SECONDS)
      );
    } catch (error) {
      // Cache is an optimization only.
    }
  }
  return session;
}

function adminLogin(input) {
  const payload = input || {};
  const suppliedPin = String(payload.pin || payload.password || payload.pms_admin_pin || '').trim();
  const configuredPin = getPmsAdminPin_();
  if (!configuredPin) {
    throw new Error('PMS PIN is not configured. Set PMS_ADMIN_PIN in Apps Script Script Properties.');
  }
  if (!suppliedPin || hashIdentifier_(suppliedPin) !== hashIdentifier_(configuredPin)) {
    throw new Error('Incorrect PMS PIN.');
  }

  const sessionToken = createPmsPinSessionToken_();
  const session = writePmsPinSession_(sessionToken);
  return {
    ok: true,
    valid: true,
    sessionToken: sessionToken,
    expiresAt: session.expiresAt,
    expiresInSeconds: PMS_PIN_SESSION_TTL_SECONDS
  };
}

function adminValidateSession(input) {
  const session = readPmsPinSession_(extractPmsSessionToken_(input));
  return {
    ok: true,
    valid: !!session,
    expiresAt: session ? session.expiresAt : '',
    expiresInSeconds: session ? Math.max(0, Math.floor((Number(session.expiresAtMs || 0) - Date.now()) / 1000)) : 0
  };
}

function adminLogout(input) {
  deletePmsPinSession_(extractPmsSessionToken_(input));
  return {
    ok: true,
    loggedOut: true
  };
}

function buildPmsAccessCacheKey_(temporaryUserKey) {
  return PMS_ACCESS_CACHE_PREFIX + ':' + String(temporaryUserKey || '').trim();
}

function readPmsAccessSession_(temporaryUserKey) {
  if (!temporaryUserKey) return null;
  try {
    const raw = CacheService.getScriptCache().get(buildPmsAccessCacheKey_(temporaryUserKey));
    return raw ? JSON.parse(raw) : null;
  } catch (error) {
    return null;
  }
}

function writePmsAccessSession_(temporaryUserKey, context) {
  if (!temporaryUserKey) return;
  const accessKey = context && context.accessKey ? String(context.accessKey || '') : '';
  CacheService.getScriptCache().put(buildPmsAccessCacheKey_(temporaryUserKey), JSON.stringify({
    authorized: true,
    via: context && context.via ? context.via : 'unknown',
    activeEmail: context && context.activeEmail ? context.activeEmail : '',
    accessKeyHash: accessKey ? hashIdentifier_(accessKey) : '',
    grantedAt: new Date().toISOString()
  }), PMS_ACCESS_SESSION_TTL_SECONDS);
}

function createPmsAccessError_() {
  const error = new Error('PMS access is restricted. Use an approved admin account or protected access link.');
  error.code = 'PMS_ACCESS_REQUIRED';
  return error;
}

function isPmsAccessError_(error) {
  return !!(error && error.code === 'PMS_ACCESS_REQUIRED');
}

function isLocalNodeHarness_() {
  return typeof process !== 'undefined' &&
    !!process &&
    !!process.versions &&
    !!process.versions.node;
}

function requirePmsAdminAccess_(input) {
  if (isLocalNodeHarness_()) {
    return { authorized: true, via: 'local-harness', activeEmail: '', accessKey: '' };
  }
  const allowlist = getPmsAdminAllowlist_();
  const configuredAccessKey = getPmsAdminAccessKey_();
  const ownerFallbackEnabled = getPmsAllowOwnerFallback_();
  const activeEmail = getPmsActiveUserEmail_();
  const effectiveEmail = getPmsEffectiveUserEmail_();
  const temporaryUserKey = getPmsTemporaryUserKey_();
  const providedAccessKey = extractPmsAccessKey_(input);
  const providedSessionToken = extractPmsSessionToken_(input);
  const pinSession = providedSessionToken ? readPmsPinSession_(providedSessionToken) : null;
  if (pinSession) {
    return {
      authorized: true,
      via: 'pin',
      activeEmail: activeEmail || String(pinSession.activeEmail || ''),
      accessKey: ''
    };
  }

  const cachedSession = temporaryUserKey ? readPmsAccessSession_(temporaryUserKey) : null;

  if (cachedSession && cachedSession.authorized) {
    const sessionVia = String(cachedSession.via || 'session').trim();
    const allowCachedSession = (
      (sessionVia === 'allowlist' && !!activeEmail && allowlist.indexOf(activeEmail) !== -1) ||
      (sessionVia === 'access-key' && !!configuredAccessKey && cachedSession.accessKeyHash === hashIdentifier_(configuredAccessKey)) ||
      (sessionVia === 'owner-fallback' && ownerFallbackEnabled && !!activeEmail && !!effectiveEmail && activeEmail === effectiveEmail)
    );
    if (allowCachedSession) {
      return {
        authorized: true,
        via: sessionVia,
        activeEmail: activeEmail,
        accessKey: configuredAccessKey && providedAccessKey === configuredAccessKey ? providedAccessKey : ''
      };
    }
  }

  if (activeEmail && allowlist.length && allowlist.indexOf(activeEmail) !== -1) {
    writePmsAccessSession_(temporaryUserKey, { via: 'allowlist', activeEmail: activeEmail });
    return { authorized: true, via: 'allowlist', activeEmail: activeEmail, accessKey: '' };
  }

  if (ownerFallbackEnabled && !allowlist.length && !configuredAccessKey && activeEmail && effectiveEmail && activeEmail === effectiveEmail) {
    writePmsAccessSession_(temporaryUserKey, { via: 'owner-fallback', activeEmail: activeEmail });
    return { authorized: true, via: 'owner-fallback', activeEmail: activeEmail, accessKey: '' };
  }

  if (configuredAccessKey && providedAccessKey && providedAccessKey === configuredAccessKey) {
    writePmsAccessSession_(temporaryUserKey, { via: 'access-key', activeEmail: activeEmail, accessKey: providedAccessKey });
    return { authorized: true, via: 'access-key', activeEmail: activeEmail, accessKey: providedAccessKey };
  }

  throw createPmsAccessError_();
}

function hashIdentifier_(value) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ''));
  return Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '').slice(0, 40);
}

function sanitizeSheetInput_(value) {
  if (value == null) return '';
  const text = String(value).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '');
  return /^[\s]*[=+\-@]/.test(text) ? "'" + text : text;
}

function parsePublicTimestamp_(value) {
  const raw = String(value || '').trim();
  if (!raw) return null;
  if (/^\d{10,}$/.test(raw)) {
    const millis = Number(raw);
    return Number.isFinite(millis) ? new Date(millis) : null;
  }
  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function getPublicRequestActorKey_(payload) {
  const input = payload || {};
  const temporaryUserKey = getPmsTemporaryUserKey_();
  const sessionId = String(input.client_session_id || input.clientSessionId || '').trim();
  const email = String(input.guest_email || input.guestEmail || '').trim().toLowerCase();
  const phone = String(input.guest_phone || input.guestPhone || '').replace(/[^\d+]/g, '');
  const checkIn = String(input.check_in || input.checkIn || '').trim();
  const checkOut = String(input.check_out || input.checkOut || '').trim();
  const productId = String(input.sellable_product_id || input.sellableProductId || input.product_id || input.productId || '').trim();
  return [temporaryUserKey, sessionId, email, phone, checkIn, checkOut, productId].filter(Boolean).join('|') || 'anonymous';
}

function getRequiredPublicClientSessionId_(payload) {
  const sessionId = String(payload && (payload.client_session_id || payload.clientSessionId) || '').trim();
  if (!sessionId || sessionId.length < 12) {
    throw new Error('Please reload the booking page and try again.');
  }
  return sessionId;
}

function enforceSimpleRateLimit_(scope, identifier, limit, windowSeconds, message) {
  const cacheKey = PUBLIC_RATE_LIMIT_CACHE_PREFIX + ':' + hashIdentifier_(scope + '|' + String(identifier || ''));
  const cache = CacheService.getScriptCache();
  const currentCount = Number(cache.get(cacheKey) || 0);
  if (currentCount >= Number(limit || 1)) {
    throw new Error(message);
  }
  cache.put(cacheKey, String(currentCount + 1), Math.max(1, Number(windowSeconds || 60)));
}

function enforcePublicSearchProtection_(payload) {
  getRequiredPublicClientSessionId_(payload);
  const searchKey = getPublicRequestActorKey_(payload);
  enforceSimpleRateLimit_(
    'search',
    searchKey + '|' + String(payload.check_in || payload.checkIn || '') + '|' + String(payload.check_out || payload.checkOut || ''),
    PUBLIC_SEARCH_MAX_REQUESTS,
    PUBLIC_SEARCH_WINDOW_SECONDS,
    'Too many availability checks were sent just now. Please wait a moment and try again.'
  );
}

function enforcePublicSubmissionProtection_(payload, scope) {
  const input = payload || {};
  getRequiredPublicClientSessionId_(input);
  const honeypot = String(input.website_url || input.websiteUrl || input.company || input.company_name || '').trim();
  if (honeypot) {
    throw new Error('We could not process that request. Please try again.');
  }

  const startedAt = parsePublicTimestamp_(input.form_started_at || input.formStartedAt);
  if (!startedAt) {
    throw new Error('Please reload the booking page and try again.');
  }
  const ageMs = Date.now() - startedAt.getTime();
  if (ageMs < PUBLIC_FORM_MIN_AGE_MS) {
    throw new Error('Please review the booking details and try again.');
  }
  if (ageMs > PUBLIC_FORM_MAX_AGE_MS) {
    throw new Error('This booking form has expired. Please refresh the page and try again.');
  }

  const actorKey = getPublicRequestActorKey_(input);
  const primaryScope = String(scope || 'request') === 'booking' ? 'website-booking' : 'website-request';
  const perHourLimit = primaryScope === 'website-booking' ? PUBLIC_BOOKING_MAX_REQUESTS : PUBLIC_REQUEST_MAX_REQUESTS;
  const duplicateKey = [
    String(input.guest_email || input.guestEmail || '').trim().toLowerCase(),
    String(input.guest_phone || input.guestPhone || '').replace(/[^\d+]/g, ''),
    String(input.check_in || input.checkIn || '').trim(),
    String(input.check_out || input.checkOut || '').trim(),
    String(input.sellable_product_id || input.sellableProductId || input.product_id || input.productId || '').trim()
  ].join('|');

  enforceSimpleRateLimit_(
    primaryScope,
    actorKey,
    perHourLimit,
    primaryScope === 'website-booking' ? PUBLIC_BOOKING_WINDOW_SECONDS : PUBLIC_REQUEST_WINDOW_SECONDS,
    'Too many booking attempts were received from this session. Please wait a few minutes and try again.'
  );

  if (duplicateKey.replace(/\|/g, '').trim()) {
    enforceSimpleRateLimit_(
      primaryScope + '-duplicate',
      duplicateKey,
      2,
      PUBLIC_DUPLICATE_WINDOW_SECONDS,
      'We already received this booking attempt very recently. Please wait a few minutes before trying again.'
    );
  }
}

function enforcePublicEmailFollowUpProtection_(payload) {
  const input = payload || {};
  getRequiredPublicClientSessionId_(input);
  const honeypot = String(input.website_url || input.websiteUrl || input.company || input.company_name || '').trim();
  if (honeypot) {
    throw new Error('We could not process that request. Please try again.');
  }

  const bookingId = String(input.booking_id || input.bookingId || '').trim();
  const guestEmail = String(input.guest_email || input.guestEmail || '').trim().toLowerCase();
  if (!bookingId || !guestEmail || !isValidEmailAddress_(guestEmail)) {
    throw new Error('Booking email follow-up details are incomplete.');
  }

  const actorKey = getPublicRequestActorKey_(input);
  enforceSimpleRateLimit_(
    'website-booking-email-followup',
    actorKey + '|' + bookingId + '|' + guestEmail,
    4,
    15 * 60,
    'Too many email confirmation follow-up attempts were received. Please contact us on WhatsApp if needed.'
  );
}

function getPublicWebsiteSettings_() {
  const contact = getWebsiteContactDetails_();
  return {
    propertyName: contact.propertyName,
    propertyAddress: contact.propertyAddress,
    phone: contact.phone,
    whatsapp: contact.whatsapp,
    email: contact.email,
    bookingCurrency: getDefaultBookingCurrency_(),
    reportingCurrency: getReportingCurrency_()
  };
}

/**
 * Public: check availability for a requested stay.
 * Accepts roomType as either a room type ID or the display name from the website.
 */
function checkAvailability(input) {
  const availabilityStartedAt = Date.now();
  logTiming_('publicAvailability:start');
  try {
  const spreadsheet = getSpreadsheet_();
  const readinessStartedAt = Date.now();
  if (!toBoolean_(input && (input._skipStructureEnsure || input.skipStructureEnsure))) {
    assertMiniPmsReady_(spreadsheet, getAvailabilityOperationSheetNames_());
  }
  logTiming_('publicAvailability:readiness', readinessStartedAt);
  const validateStartedAt = Date.now();
  const validated = validateAvailabilityInput_(input, { publicDateBounds: true });
  const roomTypeId = resolveRoomTypeId_(validated.roomType);
  const sellableProductId = String((input && (input.sellable_product_id || input.sellableProductId || input.product_id || input.productId)) || '').trim();
  logTiming_('publicAvailability:validate', validateStartedAt);
  const inventoryStartedAt = Date.now();
  const context = input && input._availabilityContext ? input._availabilityContext : null;
  const bookingsSheet = context ? null : getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  const bookingRows = context && Array.isArray(context.bookingRows)
    ? context.bookingRows
    : getBookingRowsForInventoryRecheck_(bookingsSheet, validated.checkIn, validated.checkOut, roomTypeId);
  const blockedRows = context && Array.isArray(context.blockedRows)
    ? context.blockedRows
    : getBlockedRowsForInventoryRecheck_(spreadsheet, validated.checkIn, validated.checkOut, roomTypeId);
  logTiming_('publicAvailability:readInventory', inventoryStartedAt);
  const quoteStartedAt = Date.now();
  const contextCommercialControls = context && context.commercialControlsByRoomType
    ? (context.commercialControlsByRoomType[roomTypeId] || [])
    : null;
  const contextBaseRateRows = context && context.baseRateRowsByRoomType
    ? (context.baseRateRowsByRoomType[roomTypeId] || [])
    : null;
  const contextLegacyRateRows = context && context.legacyRateRowsByRoomType
    ? (context.legacyRateRowsByRoomType[roomTypeId] || [])
    : null;
  const quoteSnapshot = buildStayAvailabilityPricingSnapshot_(validated.checkIn, validated.checkOut, roomTypeId, Number(validated.guests || 1), {
    qtyRooms: 1,
    bedSetup: validated.bedSetup,
    sellableProductId: sellableProductId,
    excludeBookingId: String((input && (input.exclude_booking_id || input.excludeBookingId)) || '').trim(),
    bookingRows: bookingRows,
    blockedRows: blockedRows,
    bookingCountMap: context && context.bookingCountMap,
    blockedCountMap: context && context.blockedCountMap,
    demandUnitsByDateRoom: context && context.demandUnitsByDateRoom,
    roomIndex: context && context.roomIndex,
    roomTypeNameMap: context && context.roomTypeNameMap,
    inventoryTotal: context && context.inventoryMap ? context.inventoryMap[roomTypeId] : null,
    physicalInventory: context && context.inventoryMap ? context.inventoryMap[roomTypeId] : null,
    commercialControls: contextCommercialControls,
    baseRateRows: contextBaseRateRows,
    legacyRateRows: contextLegacyRateRows,
    currency: context && context.bookingCurrency
  });
  logTiming_('publicAvailability:quote', quoteStartedAt);
  const stayDates = quoteSnapshot.breakdown.map(function(day) {
    return normalizeDateInput_(day.date);
  });

  if (!quoteSnapshot.hasInventoryConfigured) {
    logTiming_('publicAvailability:return', availabilityStartedAt);
    return {
      available: false,
      nights: stayDates.length,
      estimatedPrice: null,
      remainingMin: 0,
      roomTypeId: roomTypeId,
      roomTypeName: quoteSnapshot.roomTypeName,
      message: 'This room type currently has no inventory configured.',
      missingRateDates: [],
      perNight: []
    };
  }

  const perNight = quoteSnapshot.breakdown.map(function(day) {
    const resolved = quoteSnapshot.pricingByDate[day.date] || {};
    return {
      date: day.date,
      inventoryTotal: Number(day.totalRooms || quoteSnapshot.totalRooms),
      bookedConfirmed: Number(day.soldRooms || 0),
      blocked: Number(day.blockedRooms || 0),
      remaining: Number(day.availableRooms || 0),
      status: day.status || (Number(day.availableRooms || 0) < 0 ? 'Overbooked' : (Number(day.availableRooms || 0) === 0 ? 'Sold Out' : 'Available')),
      baseRate: resolved.baseRate != null ? resolved.baseRate : null,
      overrideRateUsed: resolved.overrideRateUsed != null ? resolved.overrideRateUsed : null,
      directBaseRate: resolved.directBaseRate != null ? resolved.directBaseRate : null,
      finalNightlyRate: resolved.finalNightlyRate != null ? resolved.finalNightlyRate : null,
      publicReferenceRate: resolved.publicReferenceRate != null ? resolved.publicReferenceRate : null,
      directDiscountType: resolved.directDiscountType || 'None',
      directDiscountValue: resolved.directDiscountValue === '' || resolved.directDiscountValue == null ? '' : resolved.directDiscountValue,
      savingsAmount: resolved.savingsAmount || 0,
      savingsPercentage: resolved.savingsPercentage || 0,
      pricingSource: resolved.pricingSource || 'base_rate',
      pricingReferenceId: resolved.pricingReferenceId || '',
      physicalInventory: resolved.physicalInventory != null ? resolved.physicalInventory : Number(day.totalRooms || quoteSnapshot.totalRooms),
      overbookingAllowanceApplied: resolved.overbookingAllowanceApplied != null ? resolved.overbookingAllowanceApplied : Math.max(0, Number(day.totalRooms || quoteSnapshot.totalRooms) - getRoomInventory_(roomTypeId)),
      sellableInventory: resolved.sellableInventory != null ? resolved.sellableInventory : Number(day.totalRooms || quoteSnapshot.totalRooms)
    };
  });
  const available = quoteSnapshot.availableRooms > 0 && quoteSnapshot.missingRateDates.length === 0;

  logTiming_('publicAvailability:return', availabilityStartedAt);
  return {
    available: available,
    nights: stayDates.length,
    estimatedPrice: quoteSnapshot.estimatedPrice,
    directPrice: quoteSnapshot.estimatedPrice,
    comparisonPrice: quoteSnapshot.directBookingOffer && quoteSnapshot.directBookingOffer.hasSaving ? quoteSnapshot.directBookingOffer.comparisonPrice : null,
    publicReferencePrice: quoteSnapshot.directBookingOffer && quoteSnapshot.directBookingOffer.hasSaving ? quoteSnapshot.directBookingOffer.publicReferencePrice : null,
    directDiscountType: quoteSnapshot.directBookingOffer ? quoteSnapshot.directBookingOffer.directDiscountType : 'None',
    directDiscountValue: quoteSnapshot.directBookingOffer ? quoteSnapshot.directBookingOffer.directDiscountValue : '',
    savingsAmount: quoteSnapshot.directBookingOffer && quoteSnapshot.directBookingOffer.hasSaving ? quoteSnapshot.directBookingOffer.savingsAmount : 0,
    savingsPercentage: quoteSnapshot.directBookingOffer && quoteSnapshot.directBookingOffer.hasSaving ? quoteSnapshot.directBookingOffer.savingsPercentage : 0,
    savingsLabel: quoteSnapshot.directBookingOffer && quoteSnapshot.directBookingOffer.hasSaving ? quoteSnapshot.directBookingOffer.savingsLabel : '',
    currency: (context && context.bookingCurrency) || getDefaultBookingCurrency_(),
    remainingMin: quoteSnapshot.availableRooms,
    roomTypeId: roomTypeId,
    roomTypeName: quoteSnapshot.roomTypeName,
    guests: Number(validated.guests || 1),
    bedSetup: validated.bedSetup || 'Best available',
    message: buildAvailabilityMessage_(available, quoteSnapshot.missingRateDates, quoteSnapshot.availableRooms),
    missingRateDates: quoteSnapshot.missingRateDates,
    pricingNotes: quoteSnapshot.pricingNotes,
    pricingSource: quoteSnapshot.pricingSource,
    pricingReferenceId: quoteSnapshot.pricingReferenceId,
    perNight: perNight
  };
  } finally {
    logTiming_('publicAvailability:end', availabilityStartedAt);
  }
}

function searchAvailabilityProducts(input) {
  return withRequestReadCache_(function() {
    const searchStartedAt = Date.now();
    logTiming_('publicSearch:start');
    try {
    const validated = validateStaySearchInput_(input);
    const evaluateStartedAt = Date.now();
    const products = searchSellableAvailabilityProducts_(validated.checkIn, validated.checkOut, validated.guests, validated.bedSetup);
    logTiming_('publicSearch:evaluateProducts', evaluateStartedAt);
    const bookableProducts = products.filter(function(product) {
      return product && product.bookable !== false && product.available !== false;
    });

    const response = {
      available: bookableProducts.length > 0,
      checkIn: formatDateKey_(validated.checkIn),
      checkOut: formatDateKey_(validated.checkOut),
      nights: enumerateStayDates_(validated.checkIn, validated.checkOut).length,
      guests: validated.guests,
      bedSetupRequested: validated.requestedBedSetup || 'Best available',
      bedSetupApplied: validated.bedSetup || 'Best available',
      availableProductCount: bookableProducts.length,
      unavailableProductCount: Math.max(0, products.length - bookableProducts.length),
      message: bookableProducts.length
        ? 'Select a room product to continue with your booking.'
        : 'No suitable room products are currently available for these dates and guest details.',
      products: products
    };
    logTiming_('publicSearch:return', searchStartedAt);
    return response;
    } finally {
      logTiming_('publicSearch:end', searchStartedAt);
    }
  });
}

function buildPublicAvailabilitySignal_(availableRooms) {
  const available = Math.max(0, Number(availableRooms || 0));
  if (available <= 1) {
    return {
      label: 'Last matching room',
      tone: 'last'
    };
  }
  if (available <= 2) {
    return {
      label: 'Limited availability',
      tone: 'limited'
    };
  }
  return {
    label: 'Available',
    tone: 'available'
  };
}

function buildPublicProductAvailabilityState_(snapshot) {
  const missingRateDates = snapshot && Array.isArray(snapshot.missingRateDates) ? snapshot.missingRateDates : [];
  const availableRooms = Number(snapshot && snapshot.availableRooms || 0);
  if (!snapshot || !snapshot.hasInventoryConfigured) {
    return {
      bookable: false,
      label: 'Unavailable',
      tone: 'unavailable',
      buttonLabel: 'Unavailable',
      reason: 'This room product is not available online right now.'
    };
  }
  if (snapshot.estimatedPrice == null || missingRateDates.length) {
    return {
      bookable: false,
      label: 'Unavailable',
      tone: 'unavailable',
      buttonLabel: 'Unavailable',
      reason: 'This room product cannot be priced online for these dates.'
    };
  }
  if (availableRooms < 1) {
    return {
      bookable: false,
      label: 'Sold out for these dates',
      tone: 'unavailable',
      buttonLabel: 'Sold out',
      reason: 'This room is no longer available for your selected dates.'
    };
  }
  const signal = buildPublicAvailabilitySignal_(availableRooms);
  return {
    bookable: true,
    label: signal.label,
    tone: signal.tone,
    buttonLabel: 'Book this room',
    reason: ''
  };
}

/**
 * Public: store a booking request in the Requests tab.
 */
function createRequest(input) {
  const spreadsheet = getSpreadsheet_();
  assertMiniPmsReady_(spreadsheet, getAvailabilityOperationSheetNames_().concat([SHEET_NAMES.REQUESTS]));
  const validated = validateRequestInput_(input);
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.REQUESTS);
  assertSheetHeadersReady_(sheet, REQUESTS_HEADERS);

  const now = new Date();
  const roomTypeId = resolveRoomTypeId_(validated.roomType);
  const roomTypeName = getRoomTypeNameById_(roomTypeId);
  const initialAckStatus = getInitialAcknowledgementStatus_(validated.guest_email);
  const initialAckError = getInitialAcknowledgementError_(validated.guest_email);
  const lock = LockService.getScriptLock();
  const availability = buildRequestAvailabilitySnapshot_(validated, roomTypeId, roomTypeName);

  let requestId;
  let rowNumber;

  lock.waitLock(30000);
  try {
    const headerMap = getHeaderMap_(sheet);
    requestId = generateRequestId_(sheet, headerMap, now);

    const rowObject = {
      request_id: requestId,
      created_at: now,
      guest_name: sanitizeSheetInput_(validated.guest_name),
      guest_phone: sanitizeSheetInput_(validated.guest_phone),
      guest_email: sanitizeSheetInput_(validated.guest_email),
      check_in: validated.checkIn,
      check_out: validated.checkOut,
      room_type_id: roomTypeId,
      room_type_name: roomTypeName,
      bed_setup: validated.bedSetup || 'Best available',
      guests: Number(validated.guests || 1),
      estimated_price: availability.estimatedPrice || '',
      request_status: REQUEST_STATUS_INITIAL,
      request_source: sanitizeSheetInput_(validated.request_source || REQUEST_SOURCE_WEBSITE),
      ack_email_status: initialAckStatus,
      ack_email_sent_at: '',
      ack_email_error: initialAckError,
      notes: sanitizeSheetInput_(validated.notes || ''),
      booking_id: '',
      converted_to_booking_at: '',
      conversion_status: '',
      assigned_source: '',
      assigned_booking_value: ''
    };

    rowNumber = appendObjectRow_(sheet, headerMap, rowObject);
  } finally {
    lock.releaseLock();
  }

  const acknowledgement = sendAcknowledgementEmail_({
    requestId: requestId,
    guestEmail: validated.guest_email,
    checkIn: validated.checkIn,
    checkOut: validated.checkOut,
    roomTypeName: roomTypeName,
    bedSetup: validated.bedSetup || 'Best available',
    guests: validated.guests
  });

  if (rowNumber) {
    updateObjectRow_(sheet, rowNumber, {
      ack_email_status: acknowledgement.status,
      ack_email_sent_at: acknowledgement.sentAt || '',
      ack_email_error: acknowledgement.error || ''
    });
  }

  return {
    requestId: requestId,
    requestStatus: REQUEST_STATUS_INITIAL,
    availability: availability,
    acknowledgement: acknowledgement,
    whatsappMessage: buildWhatsappMessage_({
      roomTypeName: roomTypeName,
      checkIn: validated.checkIn,
      checkOut: validated.checkOut,
      guests: validated.guests,
      bedSetup: validated.bedSetup || 'Best available'
    })
  };
}

function ensurePublicSubmissionProtectionApplied_(payload, scope) {
  const input = payload || {};
  if (input._publicSubmissionProtectionApplied) return;
  enforcePublicSubmissionProtection_(input, scope || 'request');
  input._publicSubmissionProtectionApplied = true;
}

/**
 * Public: create a real confirmed booking from the website after rechecking availability.
 */
function createWebsiteBooking(input) {
  const publicBookingStartedAt = Date.now();
  logTiming_('publicBookingSubmit:start');
  try {
  return withRequestReadCache_(function() {
    return createWebsiteBooking_(input);
  });
  } finally {
    logTiming_('publicBookingSubmit:end', publicBookingStartedAt);
  }
}

function createWebsiteBooking_(input) {
  const websiteBookingStartedAt = Date.now();
  logTiming_('websiteBooking:start');
  try {
  const validateProductStartedAt = Date.now();
  const roomRows = getRoomsMasterRows_({ activeOnly: false });
  const roomIndex = buildRoomMasterIndex_(roomRows);
  const validated = validateWebsiteBookingInput_(input, { roomIndex: roomIndex });
  const selectedProduct = validated.sellableProduct || getSellableRoomProductById_(validated.sellableProductId);
  const websiteProductConstraint = validated.sellableProductConstraint || getSellableProductConstraint_(selectedProduct, validated.bedSetup, roomIndex);
  const websiteBedSetup = websiteProductConstraint.appliedBedSetup || normalizeWebsiteBedSetupPreference_(validated.guests, validated.bedSetup);
  logTiming_('publicBookingSubmit:validate', validateProductStartedAt);
  logTiming_('websiteBooking:validateProduct', validateProductStartedAt);
  const readContextStartedAt = Date.now();
  const availabilityContext = buildPublicAvailabilitySearchContext_(validated.checkIn, validated.checkOut, [validated.roomTypeId], {
    roomRows: roomRows,
    roomIndex: roomIndex,
    timingPrefix: false
  });
  logTiming_('publicBookingSubmit:readContext', readContextStartedAt);
  const priceQuoteStartedAt = Date.now();
  const quoteRoomTypeId = validated.roomTypeId;
  const priceQuote = resolveStayCommercialQuote_(validated.checkIn, validated.checkOut, quoteRoomTypeId, Number(validated.guests || 1), {
    bookingRows: availabilityContext.bookingRows,
    blockedRows: availabilityContext.blockedRows,
    bookingCountMap: availabilityContext.bookingCountMap,
    blockedCountMap: availabilityContext.blockedCountMap,
    commercialControls: availabilityContext.commercialControlsByRoomType[quoteRoomTypeId] || [],
    baseRateRows: availabilityContext.baseRateRowsByRoomType[quoteRoomTypeId] || [],
    legacyRateRows: availabilityContext.legacyRateRowsByRoomType[quoteRoomTypeId] || [],
    physicalInventory: availabilityContext.inventoryMap[quoteRoomTypeId]
  });
  const estimatedPrice = priceQuote.total == null ? null : roundCurrency_(priceQuote.total);
  logTiming_('websiteBooking:priceQuote', priceQuoteStartedAt);

  if (estimatedPrice === null || estimatedPrice === '') {
    throw new Error('Selected stay cannot be booked online right now because pricing is unavailable. Please continue on WhatsApp.');
  }

  const bookingCurrency = availabilityContext.bookingCurrency || getDefaultBookingCurrency_();
  const fxRateToGbp = getCachedOrFallbackFxRateToGbp_(bookingCurrency);
  const finalCommitStartedAt = Date.now();
  const bookingResult = createManualBooking({
    guest_name: validated.guestName,
    guest_phone: validated.guestPhone,
    guest_email: validated.guestEmail,
    country: validated.country,
    check_in: validated.checkIn,
    check_in_time: WEBSITE_DEFAULT_CHECK_IN_TIME,
    check_out: validated.checkOut,
    check_out_time: WEBSITE_DEFAULT_CHECK_OUT_TIME,
    room_type: validated.roomTypeId,
    bed_setup: websiteBedSetup,
    guests: validated.guests,
    adults: validated.guests,
    children: 0,
    qty_rooms: 1,
    source: 'Direct Website',
    source_detail: validated.sellableProductLabel ? 'Website direct booking - ' + validated.sellableProductLabel : 'Website direct booking',
    status: BOOKING_STATUS_CONFIRMED,
    booking_value_original: estimatedPrice,
    booking_currency: bookingCurrency,
    fx_rate_to_gbp: fxRateToGbp,
    pricing_source: priceQuote.pricingSource || 'base_rate',
    pricing_reference_id: priceQuote.pricingReferenceId || '',
    _usePrecheckedLiveQuote: true,
    _precheckedLiveQuote: {
      checkIn: validated.checkIn,
      checkOut: validated.checkOut,
      roomTypeId: validated.roomTypeId,
      guests: validated.guests,
      qtyRooms: 1,
      bedSetup: websiteBedSetup,
      sellableProductId: validated.sellableProductId,
      estimatedPrice: estimatedPrice,
      pricingSource: priceQuote.pricingSource || 'base_rate',
      pricingReferenceId: priceQuote.pricingReferenceId || ''
    },
    amount_paid: 0,
    payment_status: 'Unpaid',
    notes: validated.notes,
    internal_notes: mergeOperationalNotes_(
      buildBookingRecoveryNote_([
        'Website post-commit follow-up deferred for fast guest confirmation',
        'Availability cache refresh not run synchronously'
      ]),
      buildWebsiteEmailFollowUpNote_('pending', {
        guestStatus: 'pending',
        internalStatus: 'pending',
        message: 'Guest confirmation and internal notification are pending follow-up sending.'
      })
    ),
    sellable_product_id: validated.sellableProductId,
    _skipRefresh: true,
    _skipAvailabilityRefresh: true,
    _skipStructureEnsure: true,
    _skipLiveFx: true,
    _fastInventoryRows: true,
    _deferPostCommit: true
  });
  logTiming_('websiteBooking:appendBooking', finalCommitStartedAt);

  const emailQueueStartedAt = Date.now();
  const emailFollowUp = queueWebsiteBookingEmailFollowUp_(bookingResult.bookingId);
  logTiming_('websiteBooking:emailQueue', emailQueueStartedAt);
  logTiming_('publicBookingSubmit:emailQueuedOrSent', emailQueueStartedAt);
  const responseCurrency = bookingResult.bookingCurrency || bookingCurrency || DEFAULT_CURRENCY;
  const responseReportingCurrency = bookingResult.reportingCurrency || getReportingCurrency_();
  const guestConfirmation = {
    status: emailFollowUp.status || 'pending',
    sentAt: '',
    queuedAt: emailFollowUp.queuedAt || '',
    error: emailFollowUp.error || ''
  };
  const internalNotification = {
    status: emailFollowUp.status || 'pending',
    sentAt: '',
    queuedAt: emailFollowUp.queuedAt || '',
    error: emailFollowUp.error || ''
  };

  const response = {
    ok: true,
    bookingId: bookingResult.bookingId,
    status: BOOKING_STATUS_CONFIRMED,
    source: bookingResult.source,
    roomTypeId: bookingResult.roomTypeId,
    roomTypeName: bookingResult.roomTypeName,
    checkIn: bookingResult.checkIn,
    checkOut: bookingResult.checkOut,
    nights: bookingResult.nights,
    guests: validated.guests,
    bookingValue: bookingResult.bookingValueOriginal || bookingResult.bookingValue,
    bookingValueOriginal: bookingResult.bookingValueOriginal || bookingResult.bookingValue,
    bookingValueGbp: bookingResult.bookingValueGbp || bookingResult.bookingValue,
    balanceDue: bookingResult.balanceDue,
    currency: responseCurrency,
    bookingCurrency: responseCurrency,
    reportingCurrency: responseReportingCurrency,
    pricingSource: bookingResult.pricingSource || priceQuote.pricingSource || 'base_rate',
    pricingReferenceId: bookingResult.pricingReferenceId || priceQuote.pricingReferenceId || '',
    guestConfirmation: guestConfirmation,
    internalNotification: internalNotification,
    recoveryNeeded: bookingResult.recoveryNeeded,
    recoveryMessage: bookingResult.recoveryMessage,
    postCommit: bookingResult.postCommit || null
  };
  logTiming_('publicBookingSubmit:return', websiteBookingStartedAt);
  logTiming_('websiteBooking:return', websiteBookingStartedAt);
  return response;
  } finally {
    logTiming_('websiteBooking:end', websiteBookingStartedAt);
  }
}

function saveWebsiteBookingFeedback(input) {
  const validated = validateWebsiteBookingFeedbackInput_(input);
  const bookingRef = getBookingRefById_(validated.bookingId);
  validateWebsiteBookingFeedbackBookingMatch_(validated, bookingRef.rowObject);

  const spreadsheet = getSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.WEBSITE_FEEDBACK, WEBSITE_FEEDBACK_HEADERS);
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const rowObject = {
    created_at: new Date(),
    booking_id: validated.bookingId,
    guest_name: sanitizeSheetInput_(validated.guestName),
    guest_email: sanitizeSheetInput_(validated.guestEmail),
    website_score: validated.websiteScore,
    booking_process_score: validated.bookingProcessScore,
    feedback_note: sanitizeSheetInput_(validated.feedbackNote),
    client_session_id: sanitizeSheetInput_(validated.clientSessionId),
    page_url: sanitizeSheetInput_(validated.pageUrl),
    user_agent: sanitizeSheetInput_(validated.userAgent)
  };
  const existingRowNumber = findRowNumberByHeaderValue_(sheet, 'booking_id', validated.bookingId);

  if (existingRowNumber) {
    updateObjectRowBulk_(sheet, existingRowNumber, rowObject, headers);
  } else {
    appendObjectRow_(sheet, headerMap, rowObject);
  }

  return {
    ok: true,
    saved: true,
    updated: !!existingRowNumber
  };
}

function validateWebsiteBookingFeedbackInput_(input) {
  const payload = input || {};
  const bookingId = String(payload.booking_id || payload.bookingId || '').trim();
  if (!bookingId) throw new Error('Booking ID is required for feedback.');

  const websiteScore = normalizeWebsiteFeedbackScore_(payload.website_score || payload.websiteScore);
  const bookingProcessScore = normalizeWebsiteFeedbackScore_(payload.booking_process_score || payload.bookingProcessScore);
  const feedbackNote = capString_(payload.feedback_note || payload.feedbackNote || payload.note || '', 1000);
  if (websiteScore === '' && bookingProcessScore === '' && !feedbackNote) {
    throw new Error('Please add a score or note before sending feedback.');
  }

  return {
    bookingId: capString_(bookingId, 120),
    guestName: capString_(payload.guest_name || payload.guestName || '', 200),
    guestEmail: capString_(payload.guest_email || payload.guestEmail || '', 240),
    websiteScore: websiteScore,
    bookingProcessScore: bookingProcessScore,
    feedbackNote: feedbackNote,
    clientSessionId: capString_(payload.client_session_id || payload.clientSessionId || '', 160),
    pageUrl: capString_(payload.page_url || payload.pageUrl || payload.website_url || payload.websiteUrl || '', 500),
    userAgent: capString_(payload.user_agent || payload.userAgent || '', 500)
  };
}

function normalizeWebsiteFeedbackScore_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const score = Number(raw);
  if (!isFinite(score) || score < 1 || score > 10 || Math.floor(score) !== score) {
    throw new Error('Feedback scores must be whole numbers from 1 to 10.');
  }
  return score;
}

function validateWebsiteBookingFeedbackBookingMatch_(validated, bookingRow) {
  if (!bookingRow || String(bookingRow.booking_id || '').trim() !== validated.bookingId) {
    throw new Error('Booking not found for feedback.');
  }

  const bookingEmail = String(bookingRow.guest_email || bookingRow.email || '').trim().toLowerCase();
  const feedbackEmail = String(validated.guestEmail || '').trim().toLowerCase();
  const storedSessionId = String(bookingRow.client_session_id || bookingRow.clientSessionId || '').trim();
  const feedbackSessionId = String(validated.clientSessionId || '').trim();

  if (feedbackEmail && bookingEmail && feedbackEmail !== bookingEmail) {
    throw new Error('Feedback email does not match this booking.');
  }

  if (storedSessionId && feedbackSessionId && storedSessionId !== feedbackSessionId) {
    throw new Error('Feedback session does not match this booking.');
  }

  const isDirectWebsiteBooking = isWebsiteDirectBookingSource_(bookingRow.source || bookingRow.request_source || '');
  if (isDirectWebsiteBooking) {
    if (!feedbackEmail && !isRecentBookingForWebsiteFeedback_(bookingRow)) {
      throw new Error('Feedback email is required for this booking.');
    }
    return true;
  }

  if (feedbackEmail && bookingEmail && feedbackEmail === bookingEmail) {
    return true;
  }

  if (storedSessionId && feedbackSessionId && storedSessionId === feedbackSessionId) {
    return true;
  }

  throw new Error('Feedback can only be saved for a matching website booking.');
}

function isWebsiteDirectBookingSource_(source) {
  const raw = String(source || '').trim().toLowerCase();
  const normalized = normalizeBookingSource_(source);
  return normalized === 'Direct Website' || raw.indexOf('website') !== -1 || raw.indexOf('direct') !== -1;
}

function isRecentBookingForWebsiteFeedback_(bookingRow) {
  const createdAtMs = toDateTimeMs_(bookingRow && bookingRow.created_at);
  return !!createdAtMs && Date.now() - createdAtMs <= WEBSITE_FEEDBACK_RECENT_BOOKING_WINDOW_MS;
}

function adminGetWebsiteFeedbackDashboard(input) {
  const payload = input || {};
  const limit = Math.max(1, Math.min(200, Number(payload.limit || 100) || 100));
  const spreadsheet = getSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.WEBSITE_FEEDBACK);

  if (!sheet) {
    return {
      ok: true,
      sheetAvailable: false,
      totalFeedback: 0,
      lowScoreCount: 0,
      averageWebsiteScore: null,
      averageBookingProcessScore: null,
      recentFeedback: [],
      message: 'Website_Feedback sheet is not available yet.'
    };
  }

  const sheetData = readWebsiteFeedbackSheetObjects_(sheet);
  const missingHeaders = getMissingWebsiteFeedbackDashboardHeaders_(sheetData.headers);
  const rows = sheetData.rows
    .map(normalizeWebsiteFeedbackDashboardRow_)
    .filter(function(row) {
      return row.hasDashboardContent;
    })
    .sort(function(a, b) {
      return b.createdAtMs - a.createdAtMs;
    });
  const websiteScores = rows
    .map(function(row) { return row.websiteScore; })
    .filter(function(score) { return score !== null; });
  const bookingProcessScores = rows
    .map(function(row) { return row.bookingProcessScore; })
    .filter(function(score) { return score !== null; });

  return {
    ok: true,
    sheetAvailable: true,
    totalFeedback: rows.length,
    lowScoreCount: rows.filter(function(row) { return row.isLowScore; }).length,
    averageWebsiteScore: calculateWebsiteFeedbackAverage_(websiteScores),
    averageBookingProcessScore: calculateWebsiteFeedbackAverage_(bookingProcessScores),
    recentFeedback: rows.slice(0, limit),
    limit: limit,
    generatedAt: Utilities.formatDate(new Date(), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm'),
    missingHeaders: missingHeaders,
    message: missingHeaders.length
      ? 'Website_Feedback is missing dashboard column(s): ' + missingHeaders.join(', ') + '. Showing available fields only.'
      : ''
  };
}

function readWebsiteFeedbackSheetObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values || !values.length) {
    return { headers: [], rows: [] };
  }

  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });
  if (!headers.some(function(header) { return !!header; })) {
    return { headers: [], rows: [] };
  }

  const rows = values.slice(1)
    .filter(function(row) {
      return row.some(function(cell) { return cell !== '' && cell !== null; });
    })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(header, index) {
        if (header) obj[header] = row[index];
      });
      return obj;
    });

  return {
    headers: headers.filter(function(header) { return !!header; }),
    rows: rows
  };
}

function getMissingWebsiteFeedbackDashboardHeaders_(headers) {
  const existing = {};
  (headers || []).forEach(function(header) {
    const key = String(header || '').trim();
    if (key) existing[key] = true;
  });
  return WEBSITE_FEEDBACK_DASHBOARD_HEADERS.filter(function(header) {
    return !existing[header];
  });
}

function normalizeWebsiteFeedbackDashboardRow_(row) {
  const createdAtMs = toDateTimeMs_(row.created_at);
  const websiteScore = normalizeWebsiteFeedbackDashboardScore_(row.website_score);
  const bookingProcessScore = normalizeWebsiteFeedbackDashboardScore_(row.booking_process_score);
  const isLowScore = (
    (websiteScore !== null && websiteScore <= 6) ||
    (bookingProcessScore !== null && bookingProcessScore <= 6)
  );

  return {
    createdAt: formatWebsiteFeedbackDateTime_(row.created_at),
    createdAtDisplay: formatWebsiteFeedbackDateTime_(row.created_at),
    createdAtMs: createdAtMs,
    bookingId: String(row.booking_id || '').trim(),
    guestName: String(row.guest_name || '').trim(),
    guestEmail: String(row.guest_email || '').trim(),
    websiteScore: websiteScore,
    bookingProcessScore: bookingProcessScore,
    feedbackNote: String(row.feedback_note || '').trim(),
    clientSessionId: String(row.client_session_id || '').trim(),
    pageUrl: String(row.page_url || '').trim(),
    userAgent: String(row.user_agent || '').trim(),
    isLowScore: isLowScore,
    hasDashboardContent: !!(
      createdAtMs ||
      String(row.booking_id || '').trim() ||
      String(row.guest_name || '').trim() ||
      String(row.guest_email || '').trim() ||
      websiteScore !== null ||
      bookingProcessScore !== null ||
      String(row.feedback_note || '').trim()
    )
  };
}

function normalizeWebsiteFeedbackDashboardScore_(value) {
  if (value === null || value === undefined || value === '') return null;
  const score = Number(value);
  if (!isFinite(score)) return null;
  if (score < 1 || score > 10) return null;
  return score;
}

function calculateWebsiteFeedbackAverage_(scores) {
  if (!scores || !scores.length) return null;
  const total = scores.reduce(function(sum, score) {
    return sum + Number(score || 0);
  }, 0);
  return Math.round((total / scores.length) * 10) / 10;
}

function formatWebsiteFeedbackDateTime_(value) {
  const ms = toDateTimeMs_(value);
  if (!ms) return String(value || '').trim();
  return Utilities.formatDate(new Date(ms), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm');
}

function capString_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, Number(maxLength || 1000));
}

/**
 * Public: return Settings tab as a key-value object.
 */
function getSettings() {
  if (REQUEST_READ_CACHE_ && REQUEST_READ_CACHE_.settingsLoaded) {
    return clonePlainObject_(REQUEST_READ_CACHE_.settings);
  }

  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.SETTINGS);
  const values = sheet.getDataRange().getValues();
  const result = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (!key) continue;
    result[key] = values[i][1];
  }

  if (REQUEST_READ_CACHE_) {
    REQUEST_READ_CACHE_.settings = clonePlainObject_(result);
    REQUEST_READ_CACHE_.settingsLoaded = true;
  }

  return result;
}

/**
 * Optional helper: rebuild Availability_Cache for N days ahead from today.
 */
function refreshAvailabilityCache(daysAhead) {
  const totalDays = Number(daysAhead || 90);
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const cacheSheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.AVAILABILITY_CACHE);
  const roomTypeRows = getActiveRoomTypeCatalog_();
  const today = stripTime_(new Date());
  const endExclusive = addDays_(today, totalDays);
  const bookingCountMap = buildConfirmedBookingRoomCountMap_(getSheetObjects_(SHEET_NAMES.BOOKINGS), {
    startDate: today,
    endDateExclusive: endExclusive
  }).byDateRoom;
  const blockedCountMap = buildBlockedDateQtyMap_(getSheetObjects_(SHEET_NAMES.BLOCKED_DATES), {
    startDate: today,
    endDateExclusive: endExclusive
  }).byDateRoom;
  const commercialControls = getCommercialControlRows_({ activeOnly: true });

  const headers = ['date', 'room_type_id', 'inventory_total', 'booked_confirmed', 'blocked', 'remaining', 'status'];
  const rows = [headers];

  for (let offset = 0; offset < totalDays; offset++) {
    const currentDate = new Date(today);
    currentDate.setDate(currentDate.getDate() + offset);

    roomTypeRows.forEach(function(roomType) {
      const roomTypeId = String(roomType.roomTypeId || '').trim();
      const inventoryTotal = getSellableInventoryTotalForDate_(currentDate, roomTypeId, commercialControls, Number(roomType.inventoryTotal || 0));
      const bookedConfirmed = getConfirmedBookingsCount_(currentDate, roomTypeId, bookingCountMap);
      const blocked = getBlockedDatesCount_(currentDate, roomTypeId, blockedCountMap);
      const remaining = inventoryTotal - bookedConfirmed - blocked;
      rows.push([
        formatDateKey_(currentDate),
        roomTypeId,
        inventoryTotal,
        bookedConfirmed,
        blocked,
        remaining,
        remaining < 0 ? 'Overbooked' : (remaining === 0 ? 'Sold Out' : 'Available')
      ]);
    });
  }

  cacheSheet.clearContents();
  cacheSheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  return { ok: true, rowsWritten: rows.length - 1, daysAhead: totalDays };
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Mini PMS')
      .addItem('Open Admin Panel', 'openAdminPanel')
      .addItem('Show Standalone PMS URL', 'showStandalonePmsUrl')
      .addItem('Refresh Reporting', 'menuRefreshReporting')
      .addItem('Refresh Availability', 'menuRefreshAvailability')
      .addItem('Send Pending Website Emails', 'processPendingWebsiteBookingEmails')
      .addItem('Retry Failed Website Emails', 'retryFailedWebsiteBookingEmails')
      .addItem('Test Interface', 'testAdminInterface')
      .addToUi();
  } catch (error) {
    // Ignore UI binding errors when the script is not opened from Sheets.
  }
}

function openAdminPanel() {
  const html = renderPmsApp_('sidebar').setTitle('Roza PMS Panel');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showStandalonePmsUrl() {
  const url = getStandalonePmsUrl_();
  SpreadsheetApp.getUi().alert(
    url ? 'Standalone PMS web app URL' : 'Standalone PMS web app URL unavailable',
    url || 'Deploy the script as a web app first, then reopen this menu item.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function renderPmsApp_(uiMode, accessContext) {
  let grantedAccess = accessContext && accessContext.authorized ? accessContext : null;
  if (!grantedAccess) {
    grantedAccess = hasPmsAdminPinConfigured_()
      ? { authorized: true, via: 'pin-login', activeEmail: '', accessKey: '' }
      : requirePmsAdminAccess_({});
  }
  const template = HtmlService.createTemplateFromFile('AdminPanel');
  template.uiMode = uiMode || 'web';
  template.standaloneUrl = getStandalonePmsUrl_();
  template.pmsAccessKey = grantedAccess && grantedAccess.accessKey ? grantedAccess.accessKey : '';
  return template
    .evaluate()
    .setTitle('Roza Guest House PMS');
}

function getStandalonePmsUrl_() {
  const scriptUrl = String(ScriptApp.getService().getUrl() || '').trim();
  if (scriptUrl) return scriptUrl;
  return String(PropertiesService.getScriptProperties().getProperty(PMS_STANDALONE_URL_PROPERTY_KEY) || '').trim();
}

function menuRefreshReporting() {
  const result = refreshMiniPmsReporting();
  SpreadsheetApp.getUi().alert(
    'Reporting refreshed.',
    'Booking nights: ' + result.bookingNights.rowsWritten +
    '\nDaily stats: ' + result.dailyStats.rowsWritten +
    '\nSnapshot date: ' + result.snapshot.snapshotDate,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function menuRefreshAvailability() {
  const result = refreshAvailabilityCache(180);
  SpreadsheetApp.getUi().alert(
    'Availability cache refreshed.',
    'Rows written: ' + result.rowsWritten + '\nDays ahead: ' + result.daysAhead,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function testAdminInterface() {
  const data = getAdminPanelBootstrapData();
  SpreadsheetApp.getUi().alert(
    'Mini PMS admin interface is connected.',
    'Room types loaded: ' + data.roomTypes.length + '\nCurrency: ' + data.currency,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function getAdminPanelBootstrapData(input) {
  const bootstrapStartedAt = Date.now();
  logTiming_('bootstrap:start');
  try {
  return withRequestReadCache_(function() {
    const payload = input || {};
    const uiMode = String(payload.ui_mode || payload.uiMode || '').trim().toLowerCase();
    const staffMode = uiMode === 'staff';
    if (staffMode) {
      return getStaffPanelBootstrapData_(payload);
    }

    const includeCommercialBootstrap = toBoolean_(payload.include_commercial_bootstrap || payload.includeCommercialBootstrap);
    const deferOperatorStatus = toBoolean_(payload.defer_operator_status || payload.deferOperatorStatus);
    const fastFirstRender = toBoolean_(payload.fast_first_render || payload.fastFirstRender);
    if (toBoolean_(payload.ensure_structure || payload.ensureStructure)) {
      ensureMiniPmsStructure_(getSpreadsheet_());
    }
    const today = formatDateKey_(new Date());
    const coreStartedAt = Date.now();
    const settings = getSettings();
    const defaultBookingCurrency = normalizeCurrencyCode_(settings.booking_currency || settings.currency || DEFAULT_CURRENCY) || DEFAULT_CURRENCY;
    const reportingCurrency = getReportingCurrency_();
    const roomRows = getRoomsMasterRows_({ activeOnly: false });
    const roomIndex = buildRoomMasterIndex_(roomRows);
    const roomTypeRows = getSheetObjects_(SHEET_NAMES.ROOM_TYPES);
    const activeRoomTypes = buildActiveRoomTypeCatalogFromRows_(roomTypeRows, roomRows);
    logTiming_('bootstrap:readCore', coreStartedAt);

    const bookingRowsStartedAt = Date.now();
    const dashboardBookings = fastFirstRender
      ? getFrontDeskBookingRowsForDate_(today)
      : getSheetObjects_(SHEET_NAMES.BOOKINGS);
    logTiming_('bootstrap:readDashboardBookings', bookingRowsStartedAt);
    const dashboardHistoryRows = fastFirstRender ? [] : null;
    const roomTypes = activeRoomTypes.map(function(row) {
      return {
        id: row.roomTypeId,
        name: row.roomTypeName,
        inventoryTotal: row.inventoryTotal,
        maxGuests: row.maxGuests
      };
    });
    const dashboardStartedAt = Date.now();
    const dashboard = buildFrontDeskDashboardData_(today, {
      staffMode: false,
      roomTypes: activeRoomTypes,
      bookings: dashboardBookings,
      bookingsAreOperationalSubset: fastFirstRender,
      historyRows: dashboardHistoryRows,
      roomIndex: roomIndex,
      defaultBookingCurrency: defaultBookingCurrency,
      reportingCurrency: reportingCurrency,
      fastOperational: true,
      skipGuestHistory: fastFirstRender,
      skipAvailabilitySnapshot: fastFirstRender
    });
    logTiming_('bootstrap:dashboard', dashboardStartedAt);
    if (fastFirstRender) {
      logTiming_('bootstrap:fastPayload', bootstrapStartedAt);
    }

    const bootstrap = {
      roomTypes: roomTypes,
      rooms: buildRoomMasterBootstrapRowsFromRows_(roomRows),
      bookingSourceOptions: getBookingSourceOptions_(),
      bookingStatusOptions: getBookingStatusOptions_(),
      paymentStatusOptions: getPaymentStatusOptions_(),
      paymentMethodOptions: getPaymentMethodOptions_(),
      bedSetupOptions: getBedSetupOptions_(),
      commercialRuleTypeOptions: getCommercialRuleTypeOptions_(),
      bookingCurrencyDefault: defaultBookingCurrency,
      bookingCurrencyOptions: getBookingCurrencyOptionsForDefault_(defaultBookingCurrency),
      currency: reportingCurrency,
      reportingCurrency: reportingCurrency,
      standaloneUrl: getStandalonePmsUrl_(),
      today: today,
      dashboard: dashboard
    };

    if (!deferOperatorStatus) {
      bootstrap.operatorStatus = buildPmsOperatorStatus_({
        contactDetails: getWebsiteContactDetails_(settings),
        skipMailQuota: true
      });
    } else {
      logTiming_('bootstrap:operatorStatusSkipped', bootstrapStartedAt);
    }

    if (includeCommercialBootstrap) {
      bootstrap.baseRates = buildBaseRatesData_();
      bootstrap.rateBoard = buildRateBoardData_(today);
      bootstrap.commercialControls = buildCommercialControlsData_(today);
    }

    logTiming_('bootstrap:return', bootstrapStartedAt);
    return bootstrap;
  });
  } finally {
    logTiming_('bootstrap:end', bootstrapStartedAt);
  }
}

function getStaffPanelBootstrapData_(payload) {
  const staffBootstrapStartedAt = Date.now();
  logTiming_('staffBootstrap:start');
  try {
  const today = formatDateKey_(new Date());
  const defaultBookingCurrency = DEFAULT_CURRENCY;
  const reportingCurrency = getReportingCurrency_();
  const roomRows = getRoomsMasterRows_({ activeOnly: false });
  const roomIndex = buildRoomMasterIndex_(roomRows);
  const dashboardBookings = getFrontDeskBookingRowsForDate_(today);

  return {
    roomTypes: [],
    rooms: [],
    bookingSourceOptions: [],
    bookingStatusOptions: [],
    paymentStatusOptions: [],
    paymentMethodOptions: getPaymentMethodOptions_(),
    bedSetupOptions: [],
    commercialRuleTypeOptions: [],
    bookingCurrencyDefault: defaultBookingCurrency,
    bookingCurrencyOptions: [],
    currency: reportingCurrency,
    reportingCurrency: reportingCurrency,
    standaloneUrl: getStandalonePmsUrl_(),
    today: today,
    dashboard: buildFrontDeskDashboardData_(today, {
      staffMode: true,
      bookings: dashboardBookings,
      bookingsAreOperationalSubset: true,
      roomIndex: roomIndex,
      defaultBookingCurrency: defaultBookingCurrency,
      reportingCurrency: reportingCurrency,
      fastOperational: true,
      skipGuestHistory: true
    })
  };
  } finally {
    logTiming_('staffBootstrap:end', staffBootstrapStartedAt);
  }
}

function getPmsDeferredBootstrapData(input) {
  const deferredStartedAt = Date.now();
  logTiming_('deferredBootstrap:start');
  try {
  return withRequestReadCache_(function() {
    const payload = input || {};
    const uiMode = String(payload.ui_mode || payload.uiMode || '').trim().toLowerCase();
    if (uiMode === 'staff') {
      logTiming_('deferredBootstrap:staffSkipped', deferredStartedAt);
      return {
        ok: true,
        uiMode: 'staff',
        deferred: true,
        operatorStatus: null
      };
    }

    const settings = getSettings();
    const operatorStartedAt = Date.now();
    const operatorStatus = buildPmsOperatorStatus_({
      contactDetails: getWebsiteContactDetails_(settings),
      skipMailQuota: true
    });
    logTiming_('deferredBootstrap:operatorStatus', operatorStartedAt);
    return {
      ok: true,
      uiMode: uiMode || 'full',
      deferred: true,
      operatorStatus: operatorStatus
    };
  });
  } finally {
    logTiming_('deferredBootstrap:end', deferredStartedAt);
  }
}

function refreshAvailabilityCacheWindow_(config) {
  const refreshStartedAt = Date.now();
  logTiming_('refreshAvailabilityCacheWindow_:start');
  const options = config || {};
  const startDate = normalizeDateInput_(options.startDate || options.checkIn || options.start_date || options.check_in);
  const rawEnd = normalizeDateInput_(options.endDateExclusive || options.endDate || options.checkOut || options.end_date || options.check_out);
  const inclusiveEnd = toBoolean_(options.inclusiveEnd || options.inclusive_end);

  if (!startDate || !rawEnd) {
    logTiming_('refreshAvailabilityCacheWindow:end', refreshStartedAt);
    return {
      ok: true,
      rowsWritten: 0,
      datesTouched: 0
    };
  }

  const dates = inclusiveEnd
    ? enumerateDatesInclusive_(startDate, rawEnd)
    : enumerateStayDates_(startDate, rawEnd);
  if (!dates.length) {
    logTiming_('refreshAvailabilityCacheWindow:end', refreshStartedAt);
    return {
      ok: true,
      rowsWritten: 0,
      datesTouched: 0
    };
  }

  const requestedIds = (options.roomTypeIds || options.room_types || []).map(function(roomTypeId) {
    return String(roomTypeId || '').trim();
  }).filter(Boolean);
  const spreadsheet = getSpreadsheet_();
  assertMiniPmsReady_(spreadsheet, getAvailabilityCacheOperationSheetNames_());
  const roomTypes = getActiveRoomTypeCatalog_().filter(function(roomType) {
    return !requestedIds.length || requestedIds.indexOf(roomType.roomTypeId) !== -1;
  });
  if (!roomTypes.length) {
    logTiming_('refreshAvailabilityCacheWindow:end', refreshStartedAt);
    return {
      ok: true,
      rowsWritten: 0,
      datesTouched: dates.length
    };
  }

  const cacheSheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.AVAILABILITY_CACHE);
  const readHeadersStartedAt = Date.now();
  const fallbackCacheHeaders = ['date', 'room_type_id', 'inventory_total', 'booked_confirmed', 'blocked', 'remaining', 'status'];
  const sheetLastColumn = cacheSheet.getLastColumn();
  const lastColumn = Math.max(fallbackCacheHeaders.length, sheetLastColumn);
  let headers = sheetLastColumn
    ? cacheSheet.getRange(1, 1, 1, sheetLastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  })
    : fallbackCacheHeaders.slice();
  if (!headers.some(function(header) { return !!String(header || '').trim(); })) {
    headers = fallbackCacheHeaders.slice();
  }
  logTiming_('refreshAvailabilityCacheWindow:readHeaders', readHeadersStartedAt);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const dateColumn = Object.prototype.hasOwnProperty.call(headerMap, 'date') ? headerMap.date + 1 : 1;
  const roomTypeColumn = Object.prototype.hasOwnProperty.call(headerMap, 'room_type_id') ? headerMap.room_type_id + 1 : 2;
  const lastRow = cacheSheet.getLastRow();
  const existingMap = {};
  const dateLookup = {};

  dates.forEach(function(date) {
    dateLookup[formatDateKey_(date)] = true;
  });

  if (lastRow > 1) {
    const readCacheKeysStartedAt = Date.now();
    const rowCount = lastRow - 1;
    const dateValues = cacheSheet.getRange(2, dateColumn, rowCount, 1).getValues();
    const roomTypeValues = dateColumn === roomTypeColumn
      ? dateValues
      : cacheSheet.getRange(2, roomTypeColumn, rowCount, 1).getValues();
    logTiming_('refreshAvailabilityCacheWindow:readCacheKeys', readCacheKeysStartedAt);

    for (let i = 0; i < rowCount; i++) {
      const rowDate = normalizeDateInput_(dateValues[i][0]);
      const rowRoomTypeId = String(roomTypeValues[i][0] || '').trim();
      if (!rowDate || !rowRoomTypeId) continue;
      const dateKey = formatDateKey_(rowDate);
      if (!dateLookup[dateKey]) continue;
      if (requestedIds.length && requestedIds.indexOf(rowRoomTypeId) === -1) continue;
      existingMap[dateKey + '|' + rowRoomTypeId] = i + 2;
    }
  }

  const mapBounds = {
    startDate: dates[0],
    endDateExclusive: addDays_(dates[dates.length - 1], 1)
  };
  const bookingsSheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  const readBookingsStartedAt = Date.now();
  const bookingRows = getBookingRowsForAvailabilityWindow_(bookingsSheet, mapBounds.startDate, mapBounds.endDateExclusive, requestedIds);
  logTiming_('refreshAvailabilityCacheWindow:readBookings', readBookingsStartedAt);
  const readBlockedStartedAt = Date.now();
  const blockedRows = getBlockedRowsForAvailabilityWindow_(spreadsheet, mapBounds.startDate, mapBounds.endDateExclusive, requestedIds);
  logTiming_('refreshAvailabilityCacheWindow:readBlocked', readBlockedStartedAt);
  const bookingCountMap = buildConfirmedBookingRoomCountMap_(bookingRows, mapBounds).byDateRoom;
  const blockedCountMap = buildBlockedDateQtyMap_(blockedRows, mapBounds).byDateRoom;
  const commercialControls = getCommercialControlRows_({ activeOnly: true });
  const updateRows = [];
  const appendRows = [];
  let rowsWritten = 0;

  dates.forEach(function(date) {
    const dateKey = formatDateKey_(date);
    roomTypes.forEach(function(roomType) {
      const inventoryTotal = getSellableInventoryTotalForDate_(date, roomType.roomTypeId, commercialControls, Number(roomType.inventoryTotal || 0));
      const bookedConfirmed = getConfirmedBookingsCount_(date, roomType.roomTypeId, bookingCountMap);
      const blocked = getBlockedDatesCount_(date, roomType.roomTypeId, blockedCountMap);
      const remaining = inventoryTotal - bookedConfirmed - blocked;
      const rowObject = {
        date: dateKey,
        room_type_id: roomType.roomTypeId,
        inventory_total: inventoryTotal,
        booked_confirmed: bookedConfirmed,
        blocked: blocked,
        remaining: remaining,
        status: remaining < 0 ? 'Overbooked' : (remaining === 0 ? 'Sold Out' : 'Available')
      };
      const rowValues = headers.map(function(header) {
        return Object.prototype.hasOwnProperty.call(rowObject, header) ? rowObject[header] : '';
      });
      const existingRowNumber = existingMap[dateKey + '|' + roomType.roomTypeId];

      if (existingRowNumber) {
        updateRows.push({
          rowNumber: existingRowNumber,
          values: rowValues
        });
        rowsWritten += 1;
      } else {
        appendRows.push(rowValues);
        rowsWritten += 1;
      }
    });
  });

  const writeRowsStartedAt = Date.now();
  writeRowUpdatesBatched_(cacheSheet, updateRows, headers.length);
  if (appendRows.length) {
    cacheSheet.getRange(cacheSheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
  }
  logTiming_('refreshAvailabilityCacheWindow:writeRows', writeRowsStartedAt);

  const result = {
    ok: true,
    rowsWritten: rowsWritten,
    datesTouched: dates.length,
    roomTypes: roomTypes.length
  };
  logTiming_('refreshAvailabilityCacheWindow:end', refreshStartedAt);
  return result;
}

function refreshAvailabilityCacheForTouchedBooking_(existingRow, validated) {
  const roomTypeIds = [];
  const existingRoomTypeId = String(existingRow && existingRow.room_type_id || '').trim();
  const validatedRoomTypeId = String(validated && validated.roomTypeId || '').trim();
  if (existingRoomTypeId) roomTypeIds.push(existingRoomTypeId);
  if (validatedRoomTypeId && roomTypeIds.indexOf(validatedRoomTypeId) === -1) roomTypeIds.push(validatedRoomTypeId);

  const startCandidates = [normalizeDateInput_(existingRow && existingRow.check_in), normalizeDateInput_(validated && validated.checkIn)].filter(Boolean);
  const endCandidates = [normalizeDateInput_(existingRow && existingRow.check_out), normalizeDateInput_(validated && validated.checkOut)].filter(Boolean);
  if (!startCandidates.length || !endCandidates.length || !roomTypeIds.length) {
    return {
      ok: true,
      rowsWritten: 0
    };
  }

  const startDate = startCandidates.reduce(function(min, date) {
    return date.getTime() < min.getTime() ? date : min;
  });
  const endDate = endCandidates.reduce(function(max, date) {
    return date.getTime() > max.getTime() ? date : max;
  });

  return refreshAvailabilityCacheWindow_({
    startDate: startDate,
    endDate: endDate,
    roomTypeIds: roomTypeIds
  });
}

function getCommercialControlsData(input) {
  const startedAt = Date.now();
  logTiming_('commercialControls:start');
  try {
  const payload = input || {};
  const selectedDate = normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date());
  return buildCommercialControlsData_(selectedDate);
  } finally {
    logTiming_('commercialControls:return', startedAt);
  }
}

function getRateBoardData(input) {
  const startedAt = Date.now();
  logTiming_('rateBoard:start');
  const payload = input || {};
  const selectedDate = normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date());
  const result = withRequestReadCache_(function() {
    return buildRateBoardData_(selectedDate);
  });
  logTiming_('rateBoard:return', startedAt);
  return result;
}

function getRateBoardCommercialSignals(input) {
  const startedAt = Date.now();
  logTiming_('rateBoardSignals:start');
  const payload = input || {};
  const selectedDate = normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date());
  const result = withRequestReadCache_(function() {
    return buildRateBoardCommercialSignals_(selectedDate);
  });
  logTiming_('rateBoardSignals:return', startedAt);
  return result;
}

function getRateCalendarData(input) {
  const startedAt = Date.now();
  logTiming_('rateCalendar:start');
  const validated = validateRateCalendarInput_(input);
  const result = withRequestReadCache_(function() {
    return buildRateCalendarData_(validated.startDate, validated.days, validated);
  });
  logTiming_('rateCalendar:return', startedAt);
  return result;
}

function adminSaveRateCalendarCell(input) {
  const startedAt = Date.now();
  logTiming_('rateCalendarSave:start');
  const validated = validateRateCalendarCellInput_(input);
  const spreadsheet = getSpreadsheet_();
  assertMiniPmsReady_(spreadsheet, [
    SHEET_NAMES.COMMERCIAL_CONTROLS,
    SHEET_NAMES.ROOM_TYPES
  ]);
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.COMMERCIAL_CONTROLS);
  ensureSheetHeaders_(sheet, COMMERCIAL_CONTROLS_HEADERS);
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const now = new Date();
  const findStartedAt = Date.now();
  const controls = getCommercialControlRows_({ activeOnly: false, roomTypeId: validated.roomTypeId });
  const existing = getExactSpecialCommercialControlForDate_(validated.roomTypeId, validated.date, controls);
  logTiming_('rateCalendarSave:findExisting', findStartedAt);
  const inheritedDirectDiscount = existing
    ? null
    : getRateCalendarInheritedDirectDiscountFields_(validated.roomTypeId, validated.overridePrice);
  const rowObject = {
    control_id: existing ? existing.controlId : generateCommercialControlId_(sheet, headerMap, now),
    room_type_id: validated.roomTypeId,
    room_type_name: validated.roomTypeName,
    rule_type: 'special',
    start_date: validated.date,
    end_date: validated.date,
    override_price: validated.overridePrice,
    public_reference_price: existing
      ? (existing.publicReferencePrice === '' ? '' : existing.publicReferencePrice)
      : inheritedDirectDiscount.publicReferencePrice,
    direct_discount_type: existing
      ? (existing.directDiscountType || 'None')
      : inheritedDirectDiscount.directDiscountType,
    direct_discount_value: existing
      ? (existing.directDiscountValue === '' ? '' : existing.directDiscountValue)
      : inheritedDirectDiscount.directDiscountValue,
    overbooking_allowance: validated.overbookingAllowance == null
      ? Number(existing && existing.overbookingAllowance || 0)
      : validated.overbookingAllowance,
    active: 'Yes',
    note: validated.note || (existing ? existing.note : 'Rate Calendar exact-date override'),
    created_at: existing && existing.createdAt ? existing.createdAt : now,
    updated_at: now
  };

  const writeStartedAt = Date.now();
  if (existing && existing.rowNumber) {
    updateObjectRowBulk_(sheet, existing.rowNumber, rowObject, headers);
  } else {
    appendObjectRow_(sheet, headerMap, rowObject);
  }
  logTiming_('rateCalendarSave:write', writeStartedAt);

  const cell = withRequestReadCache_(function() {
    return buildRateCalendarCellDataForDateRoom_(validated.date, validated.roomTypeId);
  });
  logTiming_('rateCalendarSave:return', startedAt);
  return {
    ok: true,
    cell: cell,
    rateBoardStale: true,
    pricingGuidanceStale: true,
    plannerStale: true,
    refreshNeeded: true
  };
}

function getBaseRatesData() {
  const startedAt = Date.now();
  logTiming_('baseRates:start');
  try {
  return buildBaseRatesData_();
  } finally {
    logTiming_('baseRates:return', startedAt);
  }
}

function adminApplyRateLadder(input) {
  const validated = validateRateLadderInput_(input);
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BASE_RATES, BASE_RATES_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  const now = new Date();
  const ladderConfig = getRateLadderConfig_();
  const baseRates = buildBaseRatesData_();
  const baseRateMap = {};
  (baseRates.rates || []).forEach(function(row) {
    baseRateMap[row.roomTypeId] = row;
  });
  const roomTypes = getRateBoardRoomTypes_();

  roomTypes.forEach(function(roomType) {
    const differential = Number(ladderConfig.differentials[roomType.roomTypeId] || 0);
    const nextBaseRate = roundCurrency_(validated.leadBaseRate + differential);
    const current = baseRateMap[roomType.roomTypeId] || {};
    const rowNumber = findRowNumberByHeaderValue_(sheet, 'room_type_id', roomType.roomTypeId);
    const rowObject = {
      room_type_id: roomType.roomTypeId,
      room_type_name: roomType.roomTypeName,
      base_rate: nextBaseRate,
      extra_guest_fee: current.extraGuestFee === '' || current.extraGuestFee == null ? '' : current.extraGuestFee,
      public_reference_price: current.publicReferencePrice === '' || current.publicReferencePrice == null ? '' : current.publicReferencePrice,
      direct_discount_type: current.directDiscountType || 'None',
      direct_discount_value: current.directDiscountValue === '' || current.directDiscountValue == null ? '' : current.directDiscountValue,
      active: current.active === false ? 'No' : 'Yes',
      updated_at: now
    };

    if (rowNumber) {
      updateObjectRow_(sheet, rowNumber, rowObject);
    } else {
      appendObjectRow_(sheet, headerMap, rowObject);
    }
  });

  saveRateLadderConfig_(Object.assign({}, ladderConfig, {
    lastReason: validated.changeReason,
    lastAppliedAt: new Date().toISOString()
  }));

  return {
    ok: true,
    leadRoomTypeId: ladderConfig.leadRoomTypeId,
    leadBaseRate: validated.leadBaseRate,
    changeReason: validated.changeReason,
    updatedCount: roomTypes.length,
    baseRates: buildBaseRatesData_(),
    rateBoard: buildRateBoardData_(validated.selectedDate)
  };
}

function adminQueueOtaFollowUpFromRateBoard(input) {
  const payload = input || {};
  const date = normalizeFrontDeskDate_(payload.date || payload.selected_date || payload.selectedDate || new Date());
  const roomTypeId = resolveRoomTypeId_(payload.room_type_id || payload.roomTypeId || payload.room_type || payload.roomType || '');
  if (!roomTypeId) throw new Error('Room type is required for OTA follow-up.');

  const recommendationAction = String(payload.recommendation_action || payload.recommendationAction || '').trim();
  const linkedActionTaken = String(payload.linked_action_taken || payload.linkedActionTaken || 'Queued from Rate Board').trim();
  const rozaPriceAfterChange = String(payload.roza_price_after_change || payload.rozaPriceAfterChange || '').trim();
  const note = String(payload.note || '').trim() || 'Queued from Rate Board for manual OTA update.';

  const saveResult = adminSaveOtaWorkflowStatus({
    date: date,
    room_type_id: roomTypeId,
    booking_com_status: 'Pending',
    airbnb_status: 'Pending',
    linked_recommendation_action: recommendationAction,
    linked_action_taken: linkedActionTaken,
    roza_price_after_change: rozaPriceAfterChange,
    note: note
  });

  return {
    ok: true,
    date: formatDateKey_(date),
    roomTypeId: roomTypeId,
    roomTypeName: getRoomTypeNameById_(roomTypeId),
    workflowData: saveResult.workflowData,
    message: 'OTA follow-up queued. Booking.com and Airbnb still need manual updating.'
  };
}

function getRecentBookingAlerts(input) {
  const alertsStartedAt = Date.now();
  logTiming_('alerts:start');
  try {
  return withRequestReadCache_(function() {
    const payload = input || {};
    const since = normalizeDateTimeInput_(payload.since || payload.last_seen_at || payload.lastSeenAt);
    const limit = Math.max(1, Math.min(12, Number(payload.limit || 6)));
    const bookingRows = getRecentBookingAlertSourceRows_();
    const rows = bookingRows
      .map(function(row) {
        return serializeRecentBookingAlertRow_(row);
      })
      .filter(function(row) {
        if (!row.createdAt) return false;
        return !since || row.createdAt.getTime() > since.getTime();
      })
      .sort(function(a, b) {
        return b.createdAt.getTime() - a.createdAt.getTime();
      })
      .slice(0, limit);

    return {
      ok: true,
      latestCreatedAt: getLatestBookingAlertCursor_(bookingRows),
      alerts: rows.map(function(row) {
        return {
          bookingId: row.bookingId,
          guestName: row.guestName,
          source: row.source,
          roomTypeName: row.roomTypeName,
          checkIn: row.checkIn,
          checkOut: row.checkOut,
          balanceDue: row.balanceDue,
          createdAt: row.createdAt ? row.createdAt.toISOString() : '',
          createdAtIso: row.createdAtIso || '',
          createdAtLabel: row.createdAtLabel,
          status: row.status,
          recoveryNeeded: !!row.recoveryNeeded,
          recoveryNote: String(row.recoveryNote || '')
        };
      }),
      operatorStatus: buildPmsOperatorStatus_({
        bookings: bookingRows,
        skipMailQuota: true
      })
    };
  });
  } finally {
    logTiming_('alerts:end', alertsStartedAt);
  }
}

function adminSaveBaseRate(input) {
  const validated = validateBaseRateInput_(input);
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BASE_RATES, BASE_RATES_HEADERS);
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const now = new Date();
  const updated = [];
  const skipped = [];

  if (validated.applyMode === 'single') {
    const rowNumber = findRowNumberByHeaderValue_(sheet, 'room_type_id', validated.roomTypeId);
    const rowObject = {
      room_type_id: validated.roomTypeId,
      room_type_name: validated.roomTypeName,
      base_rate: validated.baseRate,
      extra_guest_fee: validated.extraGuestFee === '' ? '' : validated.extraGuestFee,
      public_reference_price: validated.publicReferencePrice === '' ? '' : validated.publicReferencePrice,
      direct_discount_type: validated.directDiscountType,
      direct_discount_value: validated.directDiscountValue === '' ? '' : validated.directDiscountValue,
      active: validated.active ? 'Yes' : 'No',
      updated_at: now
    };

    if (rowNumber) {
      updateObjectRowBulk_(sheet, rowNumber, rowObject, headers);
    } else {
      appendObjectRow_(sheet, headerMap, rowObject);
    }
    updated.push({
      roomTypeId: validated.roomTypeId,
      roomTypeName: validated.roomTypeName
    });
  } else {
    const existingRows = getBaseRateRows_({ activeOnly: false });
    const existingByRoomType = existingRows.reduce(function(map, row) {
      map[row.roomTypeId] = row;
      return map;
    }, {});
    const activeRoomTypeMap = getActiveRoomTypeCatalog_().reduce(function(map, row) {
      map[row.roomTypeId] = row;
      return map;
    }, {});

    validated.targetRoomTypeIds.forEach(function(roomTypeId) {
      const roomType = activeRoomTypeMap[roomTypeId];
      if (!roomType) {
        skipped.push({
          roomTypeId: roomTypeId,
          reason: 'Room type is not active.'
        });
        return;
      }

      const existing = existingByRoomType[roomTypeId] || null;
      const fallbackRate = existing ? Number(existing.baseRate || 0) : Number(getLegacyFallbackBaseRateForRoomType_(roomTypeId) || 0);
      if (!existing && !(fallbackRate > 0)) {
        skipped.push({
          roomTypeId: roomTypeId,
          roomTypeName: roomType.roomTypeName,
          reason: 'No existing base rate was found to preserve.'
        });
        return;
      }
      try {
        validateDirectBookingDiscountFields_(validated.directDiscount, fallbackRate);
      } catch (error) {
        skipped.push({
          roomTypeId: roomTypeId,
          roomTypeName: roomType.roomTypeName,
          reason: error.message || String(error)
        });
        return;
      }

      const rowObject = {
        room_type_id: roomTypeId,
        room_type_name: roomType.roomTypeName,
        public_reference_price: validated.publicReferencePrice === '' ? '' : validated.publicReferencePrice,
        direct_discount_type: validated.directDiscountType,
        direct_discount_value: validated.directDiscountValue === '' ? '' : validated.directDiscountValue,
        updated_at: now
      };

      if (existing && existing.rowNumber) {
        updateObjectRowBulk_(sheet, existing.rowNumber, rowObject, headers);
      } else {
        const legacyFallback = getLegacyRateFallbackInfo_(roomTypeId, new Date(), getOpenRateRowsForType_(roomTypeId));
        appendObjectRow_(sheet, headerMap, Object.assign({}, rowObject, {
          base_rate: fallbackRate,
          extra_guest_fee: legacyFallback.found ? roundCurrency_(legacyFallback.extraGuestFee || 0) : '',
          active: 'Yes'
        }));
      }
      updated.push({
        roomTypeId: roomTypeId,
        roomTypeName: roomType.roomTypeName
      });
    });
  }

  return {
    ok: true,
    roomTypeId: validated.roomTypeId || '',
    applyMode: validated.applyMode,
    updatedCount: updated.length,
    updatedRoomTypes: updated,
    skipped: skipped,
    baseRates: buildBaseRatesData_()
  };
}

function adminSaveBaseRatesBulk(input) {
  const validated = validateBaseRatesBulkInput_(input);
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BASE_RATES);
  ensureSheetHeaders_(sheet, BASE_RATES_HEADERS);
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const now = new Date();
  const activeRoomTypes = getActiveRoomTypeCatalog_();
  const activeRoomTypeMap = activeRoomTypes.reduce(function(map, row) {
    map[row.roomTypeId] = row;
    return map;
  }, {});
  const existingRows = getBaseRateRows_({ activeOnly: false });
  const existingByRoomType = existingRows.reduce(function(map, row) {
    map[row.roomTypeId] = row;
    return map;
  }, {});
  const updated = [];
  const skipped = (validated.skipped || []).slice();

  validated.rows.forEach(function(row) {
    const roomType = activeRoomTypeMap[row.roomTypeId];
    if (!roomType) {
      skipped.push({
        roomTypeId: row.roomTypeId,
        reason: 'Room type is not active.'
      });
      return;
    }

    const existing = existingByRoomType[row.roomTypeId] || null;
    const rowObject = {
      room_type_id: row.roomTypeId,
      room_type_name: roomType.roomTypeName,
      base_rate: row.baseRate,
      extra_guest_fee: row.extraGuestFee === '' ? '' : row.extraGuestFee,
      public_reference_price: row.publicReferencePrice === '' ? '' : row.publicReferencePrice,
      direct_discount_type: row.directDiscountType,
      direct_discount_value: row.directDiscountValue === '' ? '' : row.directDiscountValue,
      active: row.active == null
        ? (existing ? (existing.active ? 'Yes' : 'No') : 'Yes')
        : (row.active ? 'Yes' : 'No'),
      updated_at: now
    };

    try {
      if (existing && existing.rowNumber) {
        updateObjectRowBulk_(sheet, existing.rowNumber, rowObject, headers);
      } else {
        appendObjectRow_(sheet, headerMap, rowObject);
      }
      updated.push({
        roomTypeId: row.roomTypeId,
        roomTypeName: roomType.roomTypeName
      });
    } catch (error) {
      skipped.push({
        roomTypeId: row.roomTypeId,
        roomTypeName: roomType.roomTypeName,
        reason: error.message || String(error)
      });
    }
  });

  return {
    ok: true,
    updatedCount: updated.length,
    updatedRoomTypes: updated,
    skipped: skipped,
    baseRates: buildBaseRatesData_()
  };
}

function adminSaveCommercialControl(input) {
  const commercialControlStartedAt = Date.now();
  logTiming_('commercialControl:start');
  const spreadsheet = getSpreadsheet_();
  assertMiniPmsReady_(spreadsheet, [
    SHEET_NAMES.COMMERCIAL_CONTROLS,
    SHEET_NAMES.ROOM_TYPES
  ]);
  const roomTypeNameMap = buildRoomTypeNameMap_(getSheetObjects_(SHEET_NAMES.ROOM_TYPES));
  const validated = validateCommercialControlInput_(input, {
    roomTypeNameMap: roomTypeNameMap
  });
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.COMMERCIAL_CONTROLS);
  ensureSheetHeaders_(sheet, COMMERCIAL_CONTROLS_HEADERS);
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const now = new Date();
  const rowNumber = validated.controlId
    ? findRowNumberByHeaderValue_(sheet, 'control_id', validated.controlId)
    : null;
  const existing = rowNumber ? getRowObjectByNumber_(sheet, rowNumber) : null;
  logTiming_('commercialControl:readExisting', commercialControlStartedAt);
  const controlId = rowNumber ? validated.controlId : generateCommercialControlId_(sheet, headerMap, now);
  const rowObject = {
    control_id: controlId,
    room_type_id: validated.roomTypeId,
    room_type_name: validated.roomTypeName,
    rule_type: validated.ruleType,
    start_date: validated.startDate,
    end_date: validated.endDate,
    override_price: validated.overridePrice === '' ? '' : validated.overridePrice,
    public_reference_price: validated.publicReferencePrice === '' ? '' : validated.publicReferencePrice,
    direct_discount_type: validated.directDiscountType,
    direct_discount_value: validated.directDiscountValue === '' ? '' : validated.directDiscountValue,
    overbooking_allowance: validated.overbookingAllowance,
    active: validated.active ? 'Yes' : 'No',
    note: validated.note,
    created_at: existing && existing.created_at ? existing.created_at : now,
    updated_at: now
  };

  let updateResult = null;
  if (rowNumber) {
    updateResult = updateObjectRowBulk_(sheet, rowNumber, rowObject, headers);
  } else {
    appendObjectRow_(sheet, headerMap, rowObject);
    updateResult = {
      rowNumber: sheet.getLastRow(),
      changed: true,
      updatedFields: Object.keys(rowObject)
    };
  }
  logTiming_('commercialControl:updateRow', commercialControlStartedAt);

  const inventoryControlChanged = commercialControlTouchesInventory_(existing) || commercialControlTouchesInventory_(validated);
  const cacheRefreshStatus = inventoryControlChanged
    ? {
        ok: true,
        status: 'deferred-stale',
        message: 'Availability cache refresh was deferred. Refresh the planner/Rate Board when fresh commercial output is needed.'
      }
    : {
        ok: true,
        status: 'not-needed'
      };
  logTiming_('commercialControl:postSave', commercialControlStartedAt);

  const controls = buildCommercialControlsData_(validated.selectedDate || new Date(), {
    roomTypeNameMap: roomTypeNameMap
  });
  logTiming_('commercialControl:return', commercialControlStartedAt);
  return {
    ok: true,
    controlId: controlId,
    updateResult: updateResult,
    controls: controls,
    rateBoardStale: true,
    pricingGuidanceStale: true,
    plannerStale: true,
    refreshNeeded: true,
    availabilityCacheStatus: cacheRefreshStatus.status,
    availabilityCacheUpdated: cacheRefreshStatus
  };
}

function adminSaveBulkDateRateOverrides(input) {
  const startedAt = Date.now();
  logTiming_('bulkDateRateOverride:start');
  const payload = input || {};
  const startDate = normalizeDateInput_(payload.start_date || payload.startDate);
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate);
  const rows = Array.isArray(payload.rows) ? payload.rows : [];
  if (!startDate) throw new Error('Start date is required.');
  if (!endDate) throw new Error('End date is required.');
  if (endDate.getTime() < startDate.getTime()) throw new Error('End date must be on or after start date.');
  if (!rows.length) throw new Error('Select at least one room type for a date-rate override.');

  const spreadsheet = getSpreadsheet_();
  assertMiniPmsReady_(spreadsheet, [
    SHEET_NAMES.COMMERCIAL_CONTROLS,
    SHEET_NAMES.ROOM_TYPES
  ]);
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.COMMERCIAL_CONTROLS);
  ensureSheetHeaders_(sheet, COMMERCIAL_CONTROLS_HEADERS);
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const roomTypeRows = getSheetObjects_(SHEET_NAMES.ROOM_TYPES);
  const roomTypeNameMap = buildRoomTypeNameMap_(roomTypeRows);
  const activeRoomTypeIds = getActiveRoomTypeCatalog_().map(function(roomType) {
    return roomType.roomTypeId;
  });
  const controlsStartedAt = Date.now();
  const existingControls = getCommercialControlRows_({
    activeOnly: false,
    roomTypeNameMap: roomTypeNameMap
  });
  logTiming_('bulkDateRateOverride:readCommercialControls', controlsStartedAt);

  const startKey = formatDateKey_(startDate);
  const endKey = formatDateKey_(endDate);
  const updated = [];
  const skipped = [];
  const now = new Date();

  rows.forEach(function(rawRow) {
    if (!toBoolean_(rawRow && rawRow.selected)) return;
    let validated = null;
    try {
      const roomTypeId = resolveRoomTypeId_(rawRow.room_type_id || rawRow.roomTypeId || rawRow.room_type || rawRow.roomType || '');
      if (activeRoomTypeIds.indexOf(roomTypeId) === -1) {
        throw new Error('Room type is not active.');
      }
      const existing = existingControls.filter(function(control) {
        return control.roomTypeId === roomTypeId &&
          control.startDate &&
          control.endDate &&
          formatDateKey_(control.startDate) === startKey &&
          formatDateKey_(control.endDate) === endKey;
      }).sort(compareCommercialControlPriority_)[0] || null;
      const note = String(rawRow.note || (existing && existing.note) || 'Bulk date rate override').trim();
      validated = validateCommercialControlInput_(Object.assign({}, rawRow, {
        room_type_id: roomTypeId,
        rule_type: 'special',
        start_date: startDate,
        end_date: endDate,
        active: 'Yes',
        note: note
      }), {
        roomTypeNameMap: roomTypeNameMap
      });

      const rowObject = {
        control_id: existing ? existing.controlId : generateCommercialControlId_(sheet, headerMap, now),
        room_type_id: validated.roomTypeId,
        room_type_name: validated.roomTypeName,
        rule_type: 'special',
        start_date: validated.startDate,
        end_date: validated.endDate,
        override_price: validated.overridePrice === '' ? '' : validated.overridePrice,
        public_reference_price: validated.publicReferencePrice === '' ? '' : validated.publicReferencePrice,
        direct_discount_type: validated.directDiscountType,
        direct_discount_value: validated.directDiscountValue === '' ? '' : validated.directDiscountValue,
        overbooking_allowance: existing ? Number(existing.overbookingAllowance || 0) : 0,
        active: 'Yes',
        note: validated.note,
        created_at: existing && existing.createdAt ? existing.createdAt : now,
        updated_at: now
      };

      const writeStartedAt = Date.now();
      if (existing && existing.rowNumber) {
        updateObjectRowBulk_(sheet, existing.rowNumber, rowObject, headers);
      } else {
        appendObjectRow_(sheet, headerMap, rowObject);
      }
      logTiming_('bulkDateRateOverride:write', writeStartedAt);
      updated.push({
        roomTypeId: validated.roomTypeId,
        roomTypeName: validated.roomTypeName
      });
    } catch (error) {
      skipped.push({
        roomTypeId: String(rawRow && (rawRow.room_type_id || rawRow.roomTypeId || '') || '').trim(),
        roomTypeName: validated && validated.roomTypeName ? validated.roomTypeName : '',
        reason: error.message || String(error)
      });
    }
  });

  const controls = buildCommercialControlsData_(payload.selected_date || payload.selectedDate || startDate, {
    roomTypeNameMap: roomTypeNameMap
  });
  logTiming_('bulkDateRateOverride:return', startedAt);
  return {
    ok: true,
    updatedCount: updated.length,
    updatedRoomTypes: updated,
    skipped: skipped,
    controls: controls,
    rateBoardStale: true,
    pricingGuidanceStale: true,
    plannerStale: true,
    refreshNeeded: true
  };
}

function adminSetCommercialControlActive(input) {
  const commercialControlStartedAt = Date.now();
  logTiming_('commercialControl:start');
  const payload = input || {};
  const controlId = String(payload.control_id || payload.controlId || '').trim();
  const operatorNote = String(payload.operator_note || payload.operatorNote || payload.note || payload.change_reason || payload.changeReason || '').trim();
  if (!controlId) throw new Error('Control ID is required.');

  const spreadsheet = getSpreadsheet_();
  assertMiniPmsReady_(spreadsheet, [
    SHEET_NAMES.COMMERCIAL_CONTROLS,
    SHEET_NAMES.ROOM_TYPES
  ]);
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.COMMERCIAL_CONTROLS);
  ensureSheetHeaders_(sheet, COMMERCIAL_CONTROLS_HEADERS);
  const headers = getSheetHeaders_(sheet);
  const rowNumber = findRowNumberByHeaderValue_(sheet, 'control_id', controlId);
  if (!rowNumber) throw new Error('Commercial control not found: ' + controlId);

  const existing = getRowObjectByNumber_(sheet, rowNumber);
  logTiming_('commercialControl:readExisting', commercialControlStartedAt);
  const nextActive = Object.prototype.hasOwnProperty.call(payload, 'active') || Object.prototype.hasOwnProperty.call(payload, 'isActive')
    ? toBoolean_(payload.active || payload.isActive)
    : !isYesLike_(existing.active);
  if (!operatorNote) {
    throw new Error('Add an operator reason before changing a live commercial override.');
  }
  const notePrefix = nextActive ? 'Activated override' : 'Deactivated override';

  const updateResult = updateObjectRowBulk_(sheet, rowNumber, {
    active: nextActive ? 'Yes' : 'No',
    note: mergeOperationalNotes_(existing.note, notePrefix + ': ' + operatorNote),
    updated_at: new Date()
  }, headers);
  logTiming_('commercialControl:updateRow', commercialControlStartedAt);

  const inventoryControlChanged = commercialControlTouchesInventory_(existing);
  const cacheRefreshStatus = inventoryControlChanged
    ? {
        ok: true,
        status: 'deferred-stale',
        message: 'Availability cache refresh was deferred. Refresh the planner/Rate Board when fresh commercial output is needed.'
      }
    : {
        ok: true,
        status: 'not-needed'
      };
  logTiming_('commercialControl:postSave', commercialControlStartedAt);

  const controls = buildCommercialControlsData_(payload.selected_date || payload.selectedDate || new Date(), {
    roomTypeNameMap: buildRoomTypeNameMap_(getSheetObjects_(SHEET_NAMES.ROOM_TYPES))
  });
  logTiming_('commercialControl:return', commercialControlStartedAt);
  return {
    ok: true,
    controlId: controlId,
    active: nextActive,
    updateResult: updateResult,
    controls: controls,
    rateBoardStale: true,
    pricingGuidanceStale: true,
    plannerStale: true,
    refreshNeeded: true,
    availabilityCacheStatus: cacheRefreshStatus.status,
    availabilityCacheUpdated: cacheRefreshStatus
  };
}

function commercialControlTouchesInventory_(row) {
  const source = row || {};
  return Number(source.overbookingAllowance != null ? source.overbookingAllowance : source.overbooking_allowance || 0) > 0;
}

function getEventFlagsData(input) {
  const startedAt = Date.now();
  logTiming_('eventFlags:start');
  try {
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const filters = validateEventFlagFilters_(input);
  const maxResults = Math.max(1, Math.min(100, Number((input && (input.max_results || input.maxResults)) || 50)));
  const allRows = getSheetObjects_(SHEET_NAMES.EVENT_FLAGS).map(function(row, index) {
    return normalizeEventFlagRow_(row, index + 2);
  });
  const rows = allRows.filter(function(row) {
    if (!row.startDate || !row.endDate) return false;
    if (!filters.includeInactive && !row.active) return false;
    return datesOverlapInclusive_(row.startDate, row.endDate, filters.startDate, filters.endDate);
  }).sort(compareEventFlagRows_).slice(0, maxResults);

  return {
    ok: true,
    filters: serializeEventFlagFilters_(filters),
    rows: rows.map(serializeEventFlagRow_),
    selectedDateSummary: buildEventFlagSelectedDateSummary_(filters.selectedDate, allRows)
  };
  } finally {
    logTiming_('eventFlags:return', startedAt);
  }
}

function adminSaveEventFlag(input) {
  const validated = validateEventFlagEntryInput_(input);
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.EVENT_FLAGS, EVENT_FLAGS_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  const now = new Date();
  const rowObject = {
    event_id: validated.eventId,
    event_name: validated.eventName,
    start_date: validated.startDate,
    end_date: validated.endDate,
    event_type: validated.eventType,
    impact_level: validated.impactLevel,
    note: validated.note,
    active: validated.active ? 'Yes' : 'No',
    updated_at: now
  };

  if (validated.rowNumber) {
    updateObjectRow_(sheet, validated.rowNumber, rowObject);
  } else {
    rowObject.event_id = generateEventFlagId_(sheet, headerMap, now);
    appendObjectRow_(sheet, headerMap, rowObject);
  }

  return {
    ok: true,
    eventId: rowObject.event_id,
    entry: serializeEventFlagRow_(normalizeEventFlagRow_(rowObject, validated.rowNumber || sheet.getLastRow()))
  };
}

function getCompetitorTrackerData(input) {
  const startedAt = Date.now();
  logTiming_('competitorTracker:start');
  try {
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const filters = validateCompetitorTrackerFilters_(input);
  const rows = getSheetObjects_(SHEET_NAMES.COMPETITOR_TRACKER)
    .map(function(row, index) {
      return normalizeCompetitorTrackerRow_(row, index + 2);
    })
    .filter(function(row) {
      if (!row.date) return false;
      if (row.date.getTime() < filters.startDate.getTime()) return false;
      if (row.date.getTime() > filters.endDate.getTime()) return false;
      if (filters.roomTypeId && row.matchedRozaRoomTypeId !== filters.roomTypeId) return false;
      return true;
    })
    .sort(compareCompetitorTrackerRows_);

  return {
    ok: true,
    currency: getDefaultBookingCurrency_(),
    filters: serializeCompetitorTrackerFilters_(filters),
    rows: rows.map(serializeCompetitorTrackerRow_),
    summary: buildCompetitorTrackerSummary_(rows, filters)
  };
  } finally {
    logTiming_('competitorTracker:return', startedAt);
  }
}

function adminSaveCompetitorTrackerEntry(input) {
  const validated = validateCompetitorTrackerEntryInput_(input);
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.COMPETITOR_TRACKER, COMPETITOR_TRACKER_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  const now = new Date();
  const rowObject = {
    entry_id: generateCompetitorTrackerEntryId_(sheet, headerMap, now),
    created_at: now,
    date: validated.date,
    competitor_name: validated.competitorName,
    matched_roza_room_type_id: validated.matchedRozaRoomTypeId,
    matched_roza_room_type: validated.matchedRozaRoomType,
    competitor_room_description: validated.competitorRoomDescription,
    competitor_price: validated.competitorPrice,
    available_or_sold_out: validated.availableOrSoldOut,
    breakfast_included: validated.breakfastIncluded,
    source_checked: validated.sourceChecked,
    note: validated.note
  };
  appendObjectRow_(sheet, headerMap, rowObject);

  return {
    ok: true,
    entryId: rowObject.entry_id,
    entry: serializeCompetitorTrackerRow_(normalizeCompetitorTrackerRow_(rowObject)),
    trackerData: getCompetitorTrackerData({
      start_date: validated.date,
      end_date: validated.date,
      room_type_id: validated.matchedRozaRoomTypeId
    })
  };
}

function getRecommendationActionLogData(input) {
  const startedAt = Date.now();
  logTiming_('recommendationActionLog:start');
  try {
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const filters = validateRecommendationActionLogFilters_(input);
  const maxResults = Math.max(1, Math.min(100, Number((input && (input.max_results || input.maxResults)) || 50)));
  const rows = getSheetObjects_(SHEET_NAMES.RECOMMENDATION_ACTION_LOG)
    .map(function(row, index) {
      return normalizeRecommendationActionLogRow_(row, index + 2);
    })
    .filter(function(row) {
      if (!row.date) return false;
      if (row.date.getTime() < filters.startDate.getTime()) return false;
      if (row.date.getTime() > filters.endDate.getTime()) return false;
      if (filters.roomTypeId && row.roomTypeId !== filters.roomTypeId) return false;
      return true;
    })
    .sort(compareRecommendationActionLogRows_)
    .slice(0, maxResults);

  return {
    ok: true,
    currency: getDefaultBookingCurrency_(),
    filters: serializeRecommendationActionLogFilters_(filters),
    rows: rows.map(serializeRecommendationActionLogRow_)
  };
  } finally {
    logTiming_('recommendationActionLog:return', startedAt);
  }
}

function adminSaveRecommendationActionLog(input) {
  const validated = validateRecommendationActionLogEntryInput_(input);
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.RECOMMENDATION_ACTION_LOG, RECOMMENDATION_ACTION_LOG_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  const now = new Date();
  const rowObject = {
    log_id: generateRecommendationActionLogId_(sheet, headerMap, now),
    date: validated.date,
    room_type_id: validated.roomTypeId,
    room_type_name: validated.roomTypeName,
    recommendation_action: validated.recommendationAction,
    recommendation_confidence: validated.recommendationConfidence,
    recommendation_reason_summary: validated.recommendationReasonSummary,
    roza_old_price: validated.rozaOldPrice,
    action_taken: validated.actionTaken,
    roza_new_price: validated.rozaNewPrice,
    overbooking_change: validated.overbookingChange,
    ota_updated: validated.otaUpdated,
    operator_note: validated.operatorNote,
    created_at: now
  };
  appendObjectRow_(sheet, headerMap, rowObject);

  return {
    ok: true,
    logId: rowObject.log_id,
    entry: serializeRecommendationActionLogRow_(normalizeRecommendationActionLogRow_(rowObject)),
    logData: getRecommendationActionLogData({
      start_date: validated.date,
      end_date: validated.date,
      room_type_id: validated.roomTypeId
    })
  };
}

function getOtaWorkflowData(input) {
  const startedAt = Date.now();
  logTiming_('otaWorkflow:start');
  try {
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const filters = validateOtaWorkflowFilters_(input);
  const maxResults = Math.max(1, Math.min(100, Number((input && (input.max_results || input.maxResults)) || 50)));
  const filteredRows = getSheetObjects_(SHEET_NAMES.OTA_UPDATE_WORKFLOW)
    .map(function(row, index) {
      return normalizeOtaWorkflowRow_(row, index + 2);
    })
    .filter(function(row) {
      if (!row.date) return false;
      if (row.date.getTime() < filters.startDate.getTime()) return false;
      if (row.date.getTime() > filters.endDate.getTime()) return false;
      if (filters.roomTypeId && row.roomTypeId !== filters.roomTypeId) return false;
      return true;
    });
  const rows = filteredRows.slice().sort(compareOtaWorkflowRows_).slice(0, maxResults);

  return {
    ok: true,
    currency: getDefaultBookingCurrency_(),
    filters: serializeOtaWorkflowFilters_(filters),
    rows: rows.map(serializeOtaWorkflowRow_),
    summary: buildOtaWorkflowSummary_(filteredRows, filters)
  };
  } finally {
    logTiming_('otaWorkflow:return', startedAt);
  }
}

function adminSaveOtaWorkflowStatus(input) {
  const validated = validateOtaWorkflowEntryInput_(input);
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.OTA_UPDATE_WORKFLOW, OTA_UPDATE_WORKFLOW_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  const now = new Date();
  const rowObject = {
    ota_update_id: generateOtaWorkflowUpdateId_(sheet, headerMap, now),
    date: validated.date,
    room_type_id: validated.roomTypeId,
    room_type_name: validated.roomTypeName,
    booking_com_status: validated.bookingComStatus,
    airbnb_status: validated.airbnbStatus,
    linked_recommendation_action: validated.linkedRecommendationAction,
    linked_action_taken: validated.linkedActionTaken,
    roza_price_after_change: validated.rozaPriceAfterChange == null ? '' : validated.rozaPriceAfterChange,
    note: validated.note,
    updated_at: now
  };
  appendObjectRow_(sheet, headerMap, rowObject);

  return {
    ok: true,
    otaUpdateId: rowObject.ota_update_id,
    entry: serializeOtaWorkflowRow_(normalizeOtaWorkflowRow_(rowObject)),
    workflowData: getOtaWorkflowData({
      start_date: validated.date,
      end_date: validated.date,
      room_type_id: validated.roomTypeId
    })
  };
}

function getPickupPaceData(input) {
  const startedAt = Date.now();
  logTiming_('pickupPace:start');
  try {
  ensureMiniPmsStructure_(getSpreadsheet_());
  const validated = validateDemandScoreInput_(input);
  return buildPickupPaceData_(validated.date, validated.roomTypeId);
  } finally {
    logTiming_('pickupPace:return', startedAt);
  }
}

function getDemandScoreData(input) {
  const startedAt = Date.now();
  logTiming_('demandScore:start');
  try {
  ensureMiniPmsStructure_(getSpreadsheet_());
  const validated = validateDemandScoreInput_(input);
  return buildDemandScoreData_(validated.date, validated.roomTypeId);
  } finally {
    logTiming_('demandScore:return', startedAt);
  }
}

function getPricingRecommendationData(input) {
  const startedAt = Date.now();
  logTiming_('pricingRecommendation:start');
  try {
  ensureMiniPmsStructure_(getSpreadsheet_());
  const validated = validateDemandScoreInput_(input);
  return buildPricingRecommendationData_(validated.date, validated.roomTypeId);
  } finally {
    logTiming_('pricingRecommendation:return', startedAt);
  }
}

function adminApplyManualCommercialShortcut(input) {
  ensureMiniPmsStructure_(getSpreadsheet_());
  const validated = validateManualCommercialShortcutInput_(input);
  return applyManualCommercialShortcut_(validated);
}

function adminCreateManualBooking(input) {
  const payload = input || {};

  if (!String(payload.source || payload.booking_source || '').trim()) {
    throw new Error('Booking source is required.');
  }
  if (!String(payload.guest_name || payload.guestName || '').trim()) {
    throw new Error('Guest name is required.');
  }

  return createManualBooking(payload);
}

function adminGetBookingDetail(input) {
  const bookingId = String((input && (input.booking_id || input.bookingId)) || '').trim();
  if (!bookingId) {
    throw new Error('Booking ID is required.');
  }

  return {
    ok: true,
    booking: getBookingDetailById_(bookingId)
  };
}

function adminGetRoomAssignmentOptions(input) {
  const payload = input || {};
  const bookingId = String(payload.booking_id || payload.bookingId || '').trim();
  const roomTypeId = bookingId
    ? ''
    : resolveRoomTypeId_(payload.room_type || payload.roomType || payload.room_type_id || payload.roomTypeId || '');
  const checkIn = bookingId ? null : normalizeDateInput_(payload.check_in || payload.checkIn);
  const checkOut = bookingId ? null : normalizeDateInput_(payload.check_out || payload.checkOut);

  if (!bookingId) {
    if (!roomTypeId) throw new Error('Room type is required.');
    if (!checkIn || !checkOut) throw new Error('Check-in and check-out are required.');
  }

  const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  let booking = null;
  let bookingRow = null;
  let qtyRooms = Math.max(1, Number(payload.qty_rooms || payload.qtyRooms || 1));
  let requestedGuests = Math.max(1, Number(payload.guests || payload.adults || 1));

  if (bookingId) {
    bookingRow = getBookingRefById_(bookingId).rowObject;
    const historyIndex = buildGuestHistoryIndex_(getGuestHistoryRowsForTargetRows_([bookingRow]), {
      targetBookingIds: [bookingId]
    });
    booking = mapBookingRowToAdminDisplay_(bookingRow, {
      historyIndex: historyIndex,
      roomIndex: roomIndex
    });
    qtyRooms = Math.max(1, Number(booking.qtyRooms || bookingRow.qty_rooms || 1));
    requestedGuests = Math.max(1, Number(booking.guests || bookingRow.guests || Number(bookingRow.adults || 0) + Number(bookingRow.children || 0) || 1));
  }

  const assignmentBlockedReason = qtyRooms > 1
    ? 'Room assignment currently supports one physical room per booking. Leave room assignment blank when Rooms booked is greater than 1.'
    : '';
  const assignmentOptions = assignmentBlockedReason ? [] : buildRoomAssignmentOptions_({
    roomTypeId: booking ? booking.roomTypeId : roomTypeId,
    checkIn: booking ? booking.checkIn : checkIn,
    checkOut: booking ? booking.checkOut : checkOut,
    guests: requestedGuests,
    bedSetup: booking ? booking.bedSetup : normalizeBedSetup_(payload.bed_setup || payload.bedSetup || ''),
    currentRoomIdentifier: booking ? booking.roomIdentifier : String(payload.room_identifier || payload.roomIdentifier || '').trim(),
      excludeBookingId: booking ? booking.bookingId : '',
      qtyRooms: qtyRooms
  }, getBookingRowsOverlappingStay_(booking ? bookingRow.check_in : checkIn, booking ? bookingRow.check_out : checkOut), roomIndex);

  return {
    ok: true,
    booking: booking,
    roomTypeId: booking ? booking.roomTypeId : roomTypeId,
    roomTypeName: booking ? booking.roomTypeName : getRoomTypeNameById_(roomTypeId),
    checkIn: booking ? booking.checkIn : formatDateKey_(checkIn),
    checkOut: booking ? booking.checkOut : formatDateKey_(checkOut),
    qtyRooms: qtyRooms,
    assignmentBlockedReason: assignmentBlockedReason,
    options: assignmentOptions
  };
}

function adminAssignBookingRoom(input) {
  const assignRoomStartedAt = Date.now();
  logTiming_('assignRoom:start');
  try {
  const payload = input || {};
  const bookingId = String(payload.booking_id || payload.bookingId || '').trim();
  const roomIdentifier = String(payload.room_identifier || payload.roomIdentifier || '').trim();
  const shouldRefreshDashboard = toBoolean_(payload._refreshDashboard || payload.refreshDashboard);
  const includeBookingDetail = Object.prototype.hasOwnProperty.call(payload, '_includeBookingDetail') || Object.prototype.hasOwnProperty.call(payload, 'includeBookingDetail')
    ? toBoolean_(payload._includeBookingDetail || payload.includeBookingDetail)
    : true;
  if (!bookingId) throw new Error('Booking ID is required.');

  let normalizedRoomIdentifier = '';
  const precheckRef = getBookingRefById_(bookingId);
  const precheckExisting = precheckRef.rowObject;
  const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const conflictCheckStartedAt = Date.now();
  const preloadedOverlappingRows = getBookingRowsOverlappingStay_(precheckExisting.check_in, precheckExisting.check_out);
  logTiming_('assignRoom:conflictCheck', conflictCheckStartedAt);
  const lock = LockService.getScriptLock();

  lock.waitLock(30000);
  try {
    const currentRef = getBookingRefById_(bookingId);
    const existing = currentRef.rowObject;
    const status = normalizeBookingStatus_(existing.status);
    if ([BOOKING_STATUS_CANCELLED, BOOKING_STATUS_CHECKED_OUT, BOOKING_STATUS_NO_SHOW].indexOf(status) !== -1) {
      throw new Error('Room assignment cannot be changed for ' + status + ' bookings.');
    }

    if (roomIdentifier) {
      const currentWindowMatchesPreload =
        formatDateKey_(normalizeDateInput_(existing.check_in)) === formatDateKey_(normalizeDateInput_(precheckExisting.check_in)) &&
        formatDateKey_(normalizeDateInput_(existing.check_out)) === formatDateKey_(normalizeDateInput_(precheckExisting.check_out));
      const overlappingRows = currentWindowMatchesPreload
        ? preloadedOverlappingRows
        : getBookingRowsOverlappingStay_(existing.check_in, existing.check_out);
      normalizedRoomIdentifier = normalizeRoomAssignmentForSave_(roomIdentifier, String(existing.room_type_id || '').trim(), {
        bedSetup: existing.bed_setup,
        guests: Math.max(1, Number(existing.guests || Number(existing.adults || 0) + Number(existing.children || 0) || 1)),
        qtyRooms: existing.qty_rooms,
        excludeBookingId: bookingId,
        existingRoomIdentifier: existing.room_identifier,
        checkIn: existing.check_in,
        checkOut: existing.check_out,
        roomIndex: roomIndex,
        bookingRows: overlappingRows
      });
    }

    updateObjectRowBulk_(currentRef.sheet, currentRef.rowNumber, {
      room_identifier: normalizedRoomIdentifier
    });
  } finally {
    lock.releaseLock();
  }

  if (shouldRefreshDashboard) {
    refreshMiniPmsDashboard();
  }

  return {
    ok: true,
    bookingId: bookingId,
    roomIdentifier: normalizedRoomIdentifier,
    booking: includeBookingDetail ? getBookingDetailById_(bookingId, { skipGuestHistory: true }) : null
  };
  } finally {
    logTiming_('assignRoom:end', assignRoomStartedAt);
  }
}

function adminUpdateBookingStatus(input) {
  const statusStartedAt = Date.now();
  logTiming_('status:start');
  try {
  const bookingId = String((input && (input.booking_id || input.bookingId)) || '').trim();
  const status = normalizeBookingStatus_(input && (input.status || input.booking_status));
  const shouldRefreshDashboard = toBoolean_(input && (input._refreshDashboard || input.refreshDashboard));
  const includeBookingDetail = input && (Object.prototype.hasOwnProperty.call(input, '_includeBookingDetail') || Object.prototype.hasOwnProperty.call(input, 'includeBookingDetail'))
    ? toBoolean_(input._includeBookingDetail || input.includeBookingDetail)
    : true;
  const statusChangeReason = String((input && (input.status_change_reason || input.statusChangeReason)) || '').trim();
  const allowBalanceDueCheckout = toBoolean_(input && (input.allow_balance_due_checkout || input.allowBalanceDueCheckout));
  const skipAvailabilityRefresh = toBoolean_(input && (input._skipAvailabilityRefresh || input.skipAvailabilityRefresh));

  if (!bookingId) throw new Error('Booking ID is required.');
  if (!status) throw new Error('Status is required.');

  const findStartedAt = Date.now();
  const bookingRef = getBookingRefById_(bookingId);
  const existing = bookingRef.rowObject;
  logTiming_('status:findBooking', findStartedAt);
  const existingStatus = normalizeBookingStatus_(existing.status);
  const suppliedPaymentMethod = normalizePaymentMethod_(input && (input.payment_method || input.paymentMethod) || '');
  const effectivePaymentMethod = suppliedPaymentMethod || normalizePaymentMethod_(existing.payment_method || '');
  const updatePatch = {
    status: status
  };

  if (suppliedPaymentMethod) {
    updatePatch.payment_method = suppliedPaymentMethod;
  }

  if (existingStatus === BOOKING_STATUS_CANCELLED && status !== BOOKING_STATUS_CANCELLED) {
    updatePatch.cancelled_at = '';
    updatePatch.cancel_reason = '';
  }

  if (status === BOOKING_STATUS_NO_SHOW) {
    const noShowPaymentDecision = String(input && (
      input.no_show_payment_decision ||
      input.noShowPaymentDecision ||
      input.payment_decision ||
      input.paymentDecision
    ) || statusChangeReason).trim();
    if (!noShowPaymentDecision) {
      throw new Error('Record the no-show payment decision before marking this booking No Show.');
    }
    updatePatch.internal_notes = mergeOperationalNotes_(
      existing.internal_notes,
      buildNoShowOperationalNote_(noShowPaymentDecision, getPmsOperationalOperatorLabel_())
    );
  }

  if (status === BOOKING_STATUS_IN_HOUSE) {
    requireCheckInReadiness_(Object.assign({}, existing, {
      payment_method: effectivePaymentMethod
    }));
  }

  if (status === BOOKING_STATUS_CHECKED_OUT && Number(existing.balance_due || 0) > 0.009) {
    if (!allowBalanceDueCheckout) {
      throw new Error('Outstanding balance must be cleared before check-out, or confirm the override explicitly.');
    }
    if (!statusChangeReason) {
      throw new Error('Explain why check-out is being confirmed with balance due.');
    }
    updatePatch.internal_notes = mergeOperationalNotes_(
      existing.internal_notes,
      'Check-out override with balance due ' + formatMoneyValue_(Number(existing.balance_due || 0), getReportingCurrency_()) + ': ' + statusChangeReason
    );
  }

  if (status === BOOKING_STATUS_CHECKED_OUT) {
    const actualCheckOut = normalizeDateInput_(input && (
      input.actual_check_out_date ||
      input.actualCheckOutDate ||
      input.actual_checkout_date ||
      input.actualCheckoutDate
    )) || stripTime_(new Date());
    const originalCheckIn = normalizeDateInput_(existing.check_in);
    const originalCheckOut = normalizeDateInput_(existing.check_out);
    const earlyCheckoutValueAdjusted = toBoolean_(input && (
      input.early_checkout_value_adjusted ||
      input.earlyCheckoutValueAdjusted ||
      input.value_adjusted ||
      input.valueAdjusted
    ));
    if (originalCheckIn && actualCheckOut.getTime() < stripTime_(originalCheckIn).getTime()) {
      throw new Error('Actual check-out date cannot be before check-in.');
    }
    updatePatch.actual_check_out_date = actualCheckOut;
    if (originalCheckOut && actualCheckOut.getTime() < stripTime_(originalCheckOut).getTime()) {
      updatePatch.internal_notes = mergeOperationalNotes_(
        updatePatch.internal_notes || existing.internal_notes,
        buildEarlyCheckoutOperationalNote_(originalCheckOut, actualCheckOut, earlyCheckoutValueAdjusted, getPmsOperationalOperatorLabel_())
      );
    }
    ensureSheetHeaders_(bookingRef.sheet, BOOKINGS_HEADERS);
  }

  const updateStartedAt = Date.now();
  updateObjectRowBulk_(bookingRef.sheet, bookingRef.rowNumber, updatePatch);
  logTiming_('status:updateRow', updateStartedAt);

  const availabilityStatusChanged = existingStatus !== status && (
    existingStatus === BOOKING_STATUS_CANCELLED ||
    existingStatus === BOOKING_STATUS_NO_SHOW ||
    existingStatus === BOOKING_STATUS_CHECKED_OUT ||
    status === BOOKING_STATUS_CANCELLED ||
    status === BOOKING_STATUS_NO_SHOW ||
    status === BOOKING_STATUS_CHECKED_OUT
  );

  if (!skipAvailabilityRefresh && availabilityStatusChanged) {
    const refreshStartedAt = Date.now();
    refreshAvailabilityCacheWindow_({
      startDate: existing.check_in,
      endDate: existing.check_out,
      roomTypeIds: [String(existing.room_type_id || '').trim()].filter(Boolean)
    });
    logTiming_('status:refreshAvailabilityWindow', refreshStartedAt);
  }

  if (shouldRefreshDashboard) {
    const dashboardStartedAt = Date.now();
    refreshMiniPmsDashboard();
    logTiming_('status:refreshDashboard', dashboardStartedAt);
  }

  const detailStartedAt = Date.now();
  const detailPayload = includeBookingDetail ? getBookingDetailById_(bookingId, { skipGuestHistory: true }) : null;
  logTiming_('status:detailPayload', detailStartedAt);
  logTiming_('status:return', statusStartedAt);
  return {
    ok: true,
    bookingId: bookingId,
    status: status,
    booking: detailPayload
  };
  } finally {
    logTiming_('status:end', statusStartedAt);
  }
}

function adminMarkBookingPaid(input) {
  const markPaidStartedAt = Date.now();
  logTiming_('markPaid:start');
  try {
  const payload = input || {};
  const bookingId = String(payload.booking_id || payload.bookingId || '').trim();
  const shouldRefreshDashboard = toBoolean_(payload._refreshDashboard || payload.refreshDashboard);
  const includeBookingDetail = Object.prototype.hasOwnProperty.call(payload, '_includeBookingDetail') || Object.prototype.hasOwnProperty.call(payload, 'includeBookingDetail')
    ? toBoolean_(payload._includeBookingDetail || payload.includeBookingDetail)
    : true;
  if (!bookingId) throw new Error('Booking ID is required.');

  const findStartedAt = Date.now();
  const bookingRef = getBookingRefById_(bookingId);
  const existing = bookingRef.rowObject;
  logTiming_('markPaid:findBooking', findStartedAt);
  const bookingValue = getOperationalBookingValueGbp_(existing);
  const paymentMethod = normalizePaymentMethod_(payload.payment_method || payload.paymentMethod || existing.payment_method || '');
  if (!paymentMethod) {
    throw new Error('Payment method is required to mark booking paid.');
  }
  const suppliedAmount = payload.amount_paid != null || payload.amountPaid != null || payload.amount != null
    ? validateMoneyAmountInput_(payload.amount_paid != null ? payload.amount_paid : (payload.amountPaid != null ? payload.amountPaid : payload.amount), 'Amount paid', 0)
    : bookingValue;
  if (suppliedAmount <= 0) {
    throw new Error('Amount paid must be greater than zero.');
  }
  if (suppliedAmount > bookingValue) {
    throw new Error('Amount paid cannot be greater than the booking value.');
  }
  if (suppliedAmount + 0.009 < bookingValue) {
    throw new Error('Amount paid must cover the booking value before marking paid.');
  }
  const paymentNote = String(payload.payment_notes || payload.paymentNotes || '').trim();
  const mergedNotes = paymentNote ? mergeOperationalNotes_(existing.payment_notes, paymentNote) : String(existing.payment_notes || '').trim();
  const paymentPatch = {
    amount_paid: suppliedAmount,
    payment_method: paymentMethod,
    payment_status: 'Paid',
    payment_notes: mergedNotes,
    balance_due: 0
  };

  const updateStartedAt = Date.now();
  updateObjectRowBulk_(bookingRef.sheet, bookingRef.rowNumber, paymentPatch);
  logTiming_('markPaid:updateRow', updateStartedAt);

  if (shouldRefreshDashboard) {
    const dashboardStartedAt = Date.now();
    refreshMiniPmsDashboard();
    logTiming_('markPaid:refreshDashboard', dashboardStartedAt);
  }

  const detailStartedAt = Date.now();
  const detailPayload = includeBookingDetail ? getBookingDetailById_(bookingId, { skipGuestHistory: true }) : null;
  logTiming_('markPaid:detailPayload', detailStartedAt);
  logTiming_('markPaid:return', markPaidStartedAt);
  return {
    ok: true,
    bookingId: bookingId,
    payment: {
      amountPaid: suppliedAmount,
      paymentMethod: paymentMethod,
      paymentStatus: 'Paid',
      paymentNotes: mergedNotes,
      balanceDue: 0,
      bookingValue: bookingValue,
      bookingValueGbp: bookingValue
    },
    booking: detailPayload
  };
  } finally {
    logTiming_('markPaid:end', markPaidStartedAt);
  }
}

function getOwnPayloadValue_(payload, keys) {
  for (let i = 0; i < keys.length; i++) {
    if (Object.prototype.hasOwnProperty.call(payload, keys[i])) return payload[keys[i]];
  }
  return undefined;
}

function valuesDifferForFastEdit_(currentValue, nextValue) {
  return String(currentValue == null ? '' : currentValue).trim() !== String(nextValue == null ? '' : nextValue).trim();
}

function bookingInventoryFieldsChangedForFastEdit_(existing, payload) {
  const proposed = Object.assign({}, existing || {}, payload || {});
  const existingGuests = Math.max(1, Number(existing.guests || Number(existing.adults || 0) + Number(existing.children || 0) || 1));
  const proposedGuests = Math.max(1, Number(proposed.guests || Number(proposed.adults || 0) + Number(proposed.children || 0) || 1));
  const proposedRoomTypeRaw = proposed.room_type || proposed.roomType || proposed.room_type_id || proposed.roomTypeId || '';
  let existingRoomTypeId = '';
  let proposedRoomTypeId = '';
  try {
    existingRoomTypeId = resolveRoomTypeId_(String(existing.room_type_id || existing.roomTypeId || '').trim());
    proposedRoomTypeId = resolveRoomTypeId_(String(proposedRoomTypeRaw || '').trim());
  } catch (error) {
    return true;
  }
  return (
    formatDateKey_(normalizeDateInput_(existing.check_in)) !== formatDateKey_(normalizeDateInput_(proposed.check_in || proposed.checkIn)) ||
    formatDateKey_(normalizeDateInput_(existing.check_out)) !== formatDateKey_(normalizeDateInput_(proposed.check_out || proposed.checkOut)) ||
    existingRoomTypeId !== proposedRoomTypeId ||
    normalizeBedSetup_(existing.bed_setup || 'Best available') !== normalizeBedSetup_(proposed.bed_setup || proposed.bedSetup || 'Best available') ||
    existingGuests !== proposedGuests ||
    Math.max(1, Number(existing.qty_rooms || 1)) !== Math.max(1, Number(proposed.qty_rooms || proposed.qtyRooms || 1)) ||
    normalizeBookingStatus_(existing.status) !== normalizeBookingStatus_(proposed.status || existing.status) ||
    String(existing.room_identifier || '').trim() !== String(proposed.room_identifier || proposed.roomIdentifier || '').trim()
  );
}

function bookingPaymentOrValueFieldsChangedForFastEdit_(existing, payload) {
  const proposed = Object.assign({}, existing || {}, payload || {});
  return (
    roundCurrency_(getStoredBookingOriginalValue_(existing)) !== roundCurrency_(proposed.booking_value_original || proposed.bookingValueOriginal || proposed.booking_value || proposed.bookingValue || 0) ||
    normalizeCurrencyCode_(getStoredBookingCurrency_(existing)) !== normalizeCurrencyCode_(proposed.booking_currency || proposed.bookingCurrency || proposed.currency || '') ||
    roundCurrency_(existing.amount_paid || 0) !== roundCurrency_(proposed.amount_paid || proposed.amountPaid || 0) ||
    normalizePaymentMethod_(existing.payment_method || '') !== normalizePaymentMethod_(proposed.payment_method || proposed.paymentMethod || '') ||
    normalizePaymentStatus_(existing.payment_status || '', getOperationalBookingValueGbp_(existing), Number(existing.amount_paid || 0)) !== normalizePaymentStatus_(proposed.payment_status || proposed.paymentStatus || '', getOperationalBookingValueGbp_(existing), Number(proposed.amount_paid || proposed.amountPaid || 0))
  );
}

function buildFastNonInventoryBookingEditPatch_(existing, payload) {
  if (!existing || !payload) return null;
  if (bookingInventoryFieldsChangedForFastEdit_(existing, payload)) return null;
  if (bookingPaymentOrValueFieldsChangedForFastEdit_(existing, payload)) return null;

  const patch = {};
  const fieldMap = [
    ['request_id', ['request_id', 'requestId']],
    ['source_detail', ['source_detail', 'sourceDetail']],
    ['guest_name', ['guest_name', 'guestName']],
    ['guest_phone', ['guest_phone', 'guestPhone']],
    ['guest_email', ['guest_email', 'guestEmail']],
    ['country', ['country', 'guest_country', 'guestCountry']],
    ['guest_preferences', ['guest_preferences', 'guestPreferences']],
    ['notes', ['notes']],
    ['internal_notes', ['internal_notes', 'internalNotes']],
    ['payment_notes', ['payment_notes', 'paymentNotes']]
  ];
  fieldMap.forEach(function(entry) {
    const value = getOwnPayloadValue_(payload, entry[1]);
    if (value !== undefined) patch[entry[0]] = String(value == null ? '' : value).trim();
  });
  const sourceValue = getOwnPayloadValue_(payload, ['source', 'booking_source', 'bookingSource', 'request_source', 'requestSource']);
  if (sourceValue !== undefined) patch.source = normalizeBookingSource_(sourceValue);

  const changedFields = Object.keys(patch).filter(function(key) {
    return valuesDifferForFastEdit_(existing[key], patch[key]);
  });
  if (!changedFields.length) return null;
  return {
    patch: patch,
    reportingDirty: changedFields.indexOf('source') !== -1
  };
}

function adminUpdateBooking(input) {
  const editStartedAt = Date.now();
  logTiming_('edit:start');
  try {
  const payload = input || {};
  const bookingId = String(payload.booking_id || payload.bookingId || '').trim();
  const deferHeavyRefresh = Object.prototype.hasOwnProperty.call(payload, '_skipRefresh') || Object.prototype.hasOwnProperty.call(payload, 'skipRefresh')
    ? toBoolean_(payload._skipRefresh || payload.skipRefresh)
    : !toBoolean_(payload._forceRefresh || payload.forceRefresh);
  const includeBookingDetail = Object.prototype.hasOwnProperty.call(payload, '_includeBookingDetail') || Object.prototype.hasOwnProperty.call(payload, 'includeBookingDetail')
    ? toBoolean_(payload._includeBookingDetail || payload.includeBookingDetail)
    : true;
  if (!bookingId) throw new Error('Booking ID is required.');

  let comparisonExisting = null;
  const validated = {};
  const spreadsheet = getSpreadsheet_();
  const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const lock = LockService.getScriptLock();

  const precheckRef = getBookingRefById_(bookingId);
  comparisonExisting = precheckRef.rowObject;
  const fastEdit = buildFastNonInventoryBookingEditPatch_(comparisonExisting, payload);
  if (fastEdit) {
    const fastLockWaitStartedAt = Date.now();
    lock.waitLock(30000);
    logTiming_('edit:lockWait', fastLockWaitStartedAt);
    try {
      const currentRef = getBookingRefById_(bookingId);
      const lockedFastEdit = buildFastNonInventoryBookingEditPatch_(currentRef.rowObject, payload);
      if (lockedFastEdit) {
        const fastUpdateStartedAt = Date.now();
        updateObjectRowBulk_(currentRef.sheet, currentRef.rowNumber, lockedFastEdit.patch);
        logTiming_('edit:updateRow', fastUpdateStartedAt);
        const fastDetailStartedAt = Date.now();
        const fastDetailPayload = includeBookingDetail ? getBookingDetailById_(bookingId, { skipGuestHistory: true }) : null;
        logTiming_('edit:detailPayload', fastDetailStartedAt);
        logTiming_('edit:return', editStartedAt);
        return {
          ok: true,
          bookingId: bookingId,
          booking: fastDetailPayload,
          refreshType: 'local',
          reportingDirty: lockedFastEdit.reportingDirty
        };
      }
    } finally {
      lock.releaseLock();
    }
  }

  const precheckPayload = Object.assign({}, comparisonExisting, payload);
  const validationStartedAt = Date.now();
  const precheckAdults = Math.max(0, Number(precheckPayload.adults || 0));
  const precheckChildren = Math.max(0, Number(precheckPayload.children || 0));
  const precheckGuestsInput = Number(precheckPayload.guests || 0);
  const precheckGuests = Math.max(1, precheckGuestsInput || precheckAdults + precheckChildren || 1);
  const precheckQtyRooms = Math.max(1, Number(precheckPayload.qty_rooms || precheckPayload.qtyRooms || 1));
  const precheckAvailabilityInput = validateAvailabilityInput_({
    check_in: precheckPayload.check_in || precheckPayload.checkIn,
    check_out: precheckPayload.check_out || precheckPayload.checkOut,
    room_type: precheckPayload.room_type || precheckPayload.roomType || precheckPayload.room_type_id || precheckPayload.roomTypeId,
    guests: getGuestsPerRoom_(precheckGuests, precheckQtyRooms),
    bed_setup: precheckPayload.bed_setup || precheckPayload.bedSetup || 'Best available'
  });
  const precheckRoomTypeId = resolveRoomTypeId_(precheckAvailabilityInput.roomType);
  const preloadedBookingLastRow = precheckRef.sheet.getLastRow();
  const preloadedBookingRows = getBookingRowsForInventoryRecheck_(precheckRef.sheet, precheckAvailabilityInput.checkIn, precheckAvailabilityInput.checkOut, precheckRoomTypeId);
  const preloadedBlockedRows = getBlockedRowsForInventoryRecheck_(spreadsheet, precheckAvailabilityInput.checkIn, precheckAvailabilityInput.checkOut, precheckRoomTypeId);
  logTiming_('edit:validation', validationStartedAt);

  const lockWaitStartedAt = Date.now();
  lock.waitLock(30000);
  logTiming_('edit:lockWait', lockWaitStartedAt);
  try {
    const currentRef = getBookingRefById_(bookingId);
    const currentExisting = currentRef.rowObject;
    comparisonExisting = currentExisting;
    const lockedPayload = Object.assign({}, currentExisting, payload);
    const lockedAdults = Math.max(0, Number(lockedPayload.adults || 0));
    const lockedChildren = Math.max(0, Number(lockedPayload.children || 0));
    const lockedGuestsInput = Number(lockedPayload.guests || 0);
    const lockedGuests = Math.max(1, lockedGuestsInput || lockedAdults + lockedChildren || 1);
    const lockedQtyRooms = Math.max(1, Number(lockedPayload.qty_rooms || lockedPayload.qtyRooms || 1));
    const lockedAvailabilityInput = validateAvailabilityInput_({
      check_in: lockedPayload.check_in || lockedPayload.checkIn,
      check_out: lockedPayload.check_out || lockedPayload.checkOut,
      room_type: lockedPayload.room_type || lockedPayload.roomType || lockedPayload.room_type_id || lockedPayload.roomTypeId,
      guests: getGuestsPerRoom_(lockedGuests, lockedQtyRooms),
      bed_setup: lockedPayload.bed_setup || lockedPayload.bedSetup || 'Best available'
    });
    const lockedRoomTypeId = resolveRoomTypeId_(lockedAvailabilityInput.roomType);
    const inventoryWindowMatchesPreload =
      formatDateKey_(lockedAvailabilityInput.checkIn) === formatDateKey_(precheckAvailabilityInput.checkIn) &&
      formatDateKey_(lockedAvailabilityInput.checkOut) === formatDateKey_(precheckAvailabilityInput.checkOut) &&
      lockedRoomTypeId === precheckRoomTypeId &&
      currentRef.sheet.getLastRow() === preloadedBookingLastRow;
    const lockedBookingRows = inventoryWindowMatchesPreload
      ? preloadedBookingRows
      : getBookingRowsForInventoryRecheck_(currentRef.sheet, lockedAvailabilityInput.checkIn, lockedAvailabilityInput.checkOut, lockedRoomTypeId);
    const lockedBlockedRows = inventoryWindowMatchesPreload
      ? preloadedBlockedRows
      : getBlockedRowsForInventoryRecheck_(spreadsheet, lockedAvailabilityInput.checkIn, lockedAvailabilityInput.checkOut, lockedRoomTypeId);
    const lockedValidated = validateManualBookingInput_(lockedPayload, {
      existingRow: currentExisting,
      bookingRows: lockedBookingRows,
      blockedRows: lockedBlockedRows,
      roomIndex: roomIndex
    });
    const finalRecheckStartedAt = Date.now();
    const liveAvailability = buildStayAvailabilityPricingSnapshot_(lockedValidated.checkIn, lockedValidated.checkOut, lockedValidated.roomTypeId, lockedValidated.guests, {
      qtyRooms: lockedValidated.qtyRooms,
      bedSetup: lockedValidated.bedSetup,
      excludeBookingId: bookingId,
      bookingRows: lockedBookingRows,
      blockedRows: lockedBlockedRows,
      roomIndex: roomIndex
    });
    logTiming_('edit:finalRecheck', finalRecheckStartedAt);
    const remainingMin = Number(liveAvailability.availableRooms || 0);

    if (!lockedValidated.allowOverbooking && remainingMin < lockedValidated.qtyRooms) {
      throw new Error('Only ' + Math.max(0, remainingMin) + ' room(s) remain for this stay.');
    }

    lockedValidated.roomIdentifier = normalizeRoomAssignmentForSave_(lockedValidated.roomIdentifier, lockedValidated.roomTypeId, {
      bedSetup: lockedValidated.bedSetup,
      guests: lockedValidated.guests,
      qtyRooms: lockedValidated.qtyRooms,
      excludeBookingId: bookingId,
      existingRoomIdentifier: currentExisting.room_identifier,
      checkIn: lockedValidated.checkIn,
      checkOut: lockedValidated.checkOut,
      roomIndex: roomIndex,
      bookingRows: lockedBookingRows
    });

    if (lockedValidated.status === BOOKING_STATUS_IN_HOUSE) {
      requireCheckInReadiness_({
        room_identifier: lockedValidated.roomIdentifier,
        payment_method: lockedValidated.paymentMethod
      });
    }
    if (lockedValidated.status === BOOKING_STATUS_CHECKED_OUT && Number(lockedValidated.balanceDue || 0) > 0.009) {
      throw new Error('Balance due must be zero before saving a booking as Checked Out.');
    }
    if (
      lockedValidated.status === BOOKING_STATUS_NO_SHOW &&
      normalizeBookingStatus_(currentExisting.status) !== BOOKING_STATUS_NO_SHOW &&
      !/NO-SHOW/i.test(String(lockedValidated.internalNotes || ''))
    ) {
      throw new Error('Use the No Show action so the payment decision and audit note are recorded.');
    }

    Object.keys(lockedValidated).forEach(function(key) {
      validated[key] = lockedValidated[key];
    });

    const updatePatch = {
      request_id: lockedValidated.requestId,
      source: lockedValidated.source,
      source_detail: lockedValidated.sourceDetail,
      guest_name: lockedValidated.guestName,
      guest_phone: lockedValidated.guestPhone,
      guest_email: lockedValidated.guestEmail,
      country: lockedValidated.country,
      check_in: lockedValidated.checkIn,
      check_in_time: lockedValidated.checkInTime,
      check_out: lockedValidated.checkOut,
      check_out_time: lockedValidated.checkOutTime,
      nights: lockedValidated.nights,
      room_type_id: lockedValidated.roomTypeId,
      room_type_name: lockedValidated.roomTypeName,
      room_identifier: lockedValidated.roomIdentifier,
      bed_setup: lockedValidated.bedSetup,
      adults: lockedValidated.adults,
      children: lockedValidated.children,
      guests: lockedValidated.guests,
      qty_rooms: lockedValidated.qtyRooms,
      status: lockedValidated.status,
      booking_value: lockedValidated.bookingValue,
      booking_value_original: lockedValidated.bookingValueOriginal,
      booking_currency: lockedValidated.bookingCurrency,
      fx_rate_to_gbp: lockedValidated.fxRateToGbp,
      booking_value_gbp: lockedValidated.bookingValueGbp,
      pricing_source: lockedValidated.pricingSource,
      pricing_reference_id: lockedValidated.pricingReferenceId,
      amount_paid: lockedValidated.amountPaid,
      payment_method: lockedValidated.paymentMethod,
      tax_amount: lockedValidated.taxAmount,
      payment_status: lockedValidated.paymentStatus,
      payment_notes: lockedValidated.paymentNotes,
      balance_due: lockedValidated.balanceDue,
      currency: lockedValidated.currency,
      cancelled_at: lockedValidated.status === BOOKING_STATUS_CANCELLED ? (currentExisting.cancelled_at || new Date()) : '',
      cancel_reason: lockedValidated.status === BOOKING_STATUS_CANCELLED ? String(currentExisting.cancel_reason || 'Cancelled manually').trim() : '',
      guest_preferences: lockedValidated.guestPreferences,
      notes: lockedValidated.notes,
      internal_notes: lockedValidated.internalNotes
    };

    const updateRowStartedAt = Date.now();
    updateObjectRowBulk_(currentRef.sheet, currentRef.rowNumber, updatePatch);
    logTiming_('edit:updateRow', updateRowStartedAt);
  } finally {
    lock.releaseLock();
  }

  const classifyStartedAt = Date.now();
  const needsFullRefresh =
    formatDateKey_(normalizeDateInput_(comparisonExisting.check_in)) !== formatDateKey_(validated.checkIn) ||
    formatDateKey_(normalizeDateInput_(comparisonExisting.check_out)) !== formatDateKey_(validated.checkOut) ||
    String(comparisonExisting.room_type_id || '').trim() !== validated.roomTypeId ||
    Number(comparisonExisting.qty_rooms || 1) !== validated.qtyRooms ||
    normalizeBookingStatus_(comparisonExisting.status) !== validated.status ||
    normalizeBookingSource_(comparisonExisting.source || '') !== validated.source ||
    roundCurrency_(getOperationalBookingValueGbp_(comparisonExisting)) !== validated.bookingValue ||
    normalizeCurrencyCode_(getStoredBookingCurrency_(comparisonExisting)) !== validated.bookingCurrency;
  logTiming_('edit:classifyChange', classifyStartedAt);

  let refreshType = 'local';
  if (needsFullRefresh && !deferHeavyRefresh) {
    refreshMiniPmsReporting();
    refreshType = 'reporting';
  } else if (needsFullRefresh) {
    refreshAvailabilityCacheForTouchedBooking_(comparisonExisting, validated);
    refreshType = 'availability-window';
  } else if (!deferHeavyRefresh) {
    refreshMiniPmsDashboard();
    refreshType = 'dashboard';
  }

  const detailPayloadStartedAt = Date.now();
  const detailPayload = includeBookingDetail ? getBookingDetailById_(bookingId, { skipGuestHistory: true }) : null;
  logTiming_('edit:detailPayload', detailPayloadStartedAt);
  logTiming_('edit:return', editStartedAt);
  return {
    ok: true,
    bookingId: bookingId,
    booking: detailPayload,
    refreshType: refreshType,
    reportingDirty: needsFullRefresh && deferHeavyRefresh
  };
  } finally {
    logTiming_('edit:end', editStartedAt);
  }
}

function adminCancelBooking(input) {
  const adminCancelStartedAt = Date.now();
  logTiming_('adminCancel:start');
  try {
  const bookingId = String((input && (input.booking_id || input.bookingId)) || '').trim();
  const reason = String((input && (input.reason || input.cancel_reason)) || 'Cancelled from admin panel').trim();
  const skipRefresh = toBoolean_(input && (input._skipRefresh || input.skipRefresh));
  const includeBookingDetail = input && (Object.prototype.hasOwnProperty.call(input, '_includeBookingDetail') || Object.prototype.hasOwnProperty.call(input, 'includeBookingDetail'))
    ? toBoolean_(input._includeBookingDetail || input.includeBookingDetail)
    : false;
  return cancelBooking(bookingId, reason, {
    skipRefresh: skipRefresh,
    includeBookingDetail: includeBookingDetail
  });
  } finally {
    logTiming_('adminCancel:end', adminCancelStartedAt);
  }
}

function adminCheckAvailability(input) {
  const payload = input || {};
  const checkIn = normalizeDateInput_(payload.check_in || payload.checkIn);
  const checkOut = normalizeDateInput_(payload.check_out || payload.checkOut);
  const roomTypeRaw = String(payload.room_type || payload.roomType || payload.room_type_id || payload.roomTypeId || '').trim();
  const qtyRooms = Math.max(1, Number(payload.qty_rooms || payload.qtyRooms || 1));
  const bedSetup = normalizeBedSetup_(payload.bed_setup || payload.bedSetup || 'Best available');

  if (!checkIn) throw new Error('Check-in date is required.');
  if (!checkOut) throw new Error('Check-out date is required.');
  if (!roomTypeRaw) throw new Error('Room type is required.');
  if (checkOut.getTime() <= checkIn.getTime()) throw new Error('Check-out must be after check-in.');

  const roomTypeId = resolveRoomTypeId_(roomTypeRaw);
  const totalGuests = Math.max(1, Number(payload.guests || payload.adults || 1));
  const guestsPerRoom = getGuestsPerRoom_(totalGuests, qtyRooms);
  if (guestsPerRoom === 3 && roomTypeId !== 'COTTAGE' && ['Double', 'Twin'].indexOf(bedSetup || '') !== -1) {
    throw new Error('For 3-guest stays, use Best available or Triple / Family.');
  }
  if (guestsPerRoom > 3 && bedSetup === 'Triple') {
    throw new Error('Triple / Family setup is not valid for this occupancy. Use Best available.');
  }
  const quoteSnapshot = buildStayAvailabilityPricingSnapshot_(checkIn, checkOut, roomTypeId, Number(payload.guests || payload.adults || 1), {
    qtyRooms: qtyRooms,
    bedSetup: bedSetup,
    excludeBookingId: String(payload.exclude_booking_id || payload.excludeBookingId || '').trim()
  });

  return {
    roomTypeId: roomTypeId,
    roomTypeName: quoteSnapshot.roomTypeName,
    checkIn: formatDateKey_(checkIn),
    checkOut: formatDateKey_(checkOut),
    nights: quoteSnapshot.breakdown.length,
    bedSetup: bedSetup || 'Best available',
    qtyRooms: qtyRooms,
    totalRooms: quoteSnapshot.totalRooms,
    soldRooms: quoteSnapshot.soldRooms,
    blockedRooms: quoteSnapshot.blockedRooms,
    availableRooms: quoteSnapshot.availableRooms,
    available: quoteSnapshot.availableRooms >= qtyRooms,
    estimatedPrice: quoteSnapshot.estimatedPrice,
    estimatedPricePerRoom: quoteSnapshot.estimatedPricePerRoom,
    currency: getDefaultBookingCurrency_(),
    pricingSource: quoteSnapshot.pricingSource,
    pricingReferenceId: quoteSnapshot.pricingReferenceId,
    pricingNotes: quoteSnapshot.pricingNotes,
    source: quoteSnapshot.source,
    dailyBreakdown: quoteSnapshot.breakdown.map(function(day) {
      const resolved = quoteSnapshot.pricingByDate[day.date] || {};
      return Object.assign({}, day, {
        baseRate: resolved.baseRate != null ? resolved.baseRate : null,
        overrideRateUsed: resolved.overrideRateUsed != null ? resolved.overrideRateUsed : null,
        finalNightlyRate: resolved.finalNightlyRate != null ? resolved.finalNightlyRate : null,
        pricingSource: resolved.pricingSource || 'base_rate',
        pricingReferenceId: resolved.pricingReferenceId || '',
        physicalInventory: resolved.physicalInventory != null ? resolved.physicalInventory : Number(day.totalRooms || quoteSnapshot.totalRooms),
        overbookingAllowanceApplied: resolved.overbookingAllowanceApplied != null ? resolved.overbookingAllowanceApplied : 0,
        sellableInventory: resolved.sellableInventory != null ? resolved.sellableInventory : Number(day.totalRooms || quoteSnapshot.totalRooms)
      });
    })
  };
}

function adminAddBlock(input) {
  const payload = input || {};
  const validated = validateBlockInput_(input);
  const spreadsheet = getSpreadsheet_();
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BLOCKED_DATES, BLOCKED_DATES_HEADERS);
  const now = new Date();
  const lock = LockService.getScriptLock();

  let blockedId;
  let rowsWritten = 0;

  lock.waitLock(30000);
  try {
    const headerMap = getHeaderMap_(sheet);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(header) {
      return String(header || '').trim();
    });
    const dates = enumerateDatesInclusive_(validated.startDate, validated.endDate);
    blockedId = generateBlockedId_(sheet, headerMap, now);

    const valueRows = dates.map(function(date) {
      const rowObject = {
        blocked_id: blockedId,
        date: date,
        room_type_id: validated.roomTypeId,
        qty_blocked: validated.qtyBlocked,
        reason: validated.reason,
        status: 'Active',
        notes: validated.notes
      };
      return headers.map(function(header) {
        return Object.prototype.hasOwnProperty.call(rowObject, header) ? rowObject[header] : '';
      });
    });

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, valueRows.length, headers.length).setValues(valueRows);
    rowsWritten = valueRows.length;
  } finally {
    lock.releaseLock();
  }

  if (toBoolean_(payload._skipRefresh || payload.skipRefresh)) {
    refreshAvailabilityCacheWindow_({
      startDate: validated.startDate,
      endDate: validated.endDate,
      roomTypeIds: [validated.roomTypeId],
      inclusiveEnd: true
    });
  } else {
    refreshMiniPmsReporting();
  }

  return {
    ok: true,
    blockedId: blockedId,
    rowsWritten: rowsWritten,
    roomTypeId: validated.roomTypeId,
    roomTypeName: validated.roomTypeName
  };
}

function adminSearchBookings(input) {
  const payload = input || {};
  const bookingId = String(payload.booking_id || payload.bookingId || '').trim().toLowerCase();
  const guestName = String(payload.guest_name || payload.guestName || '').trim().toLowerCase();
  const source = normalizeBookingSource_(payload.source || '');
  const targetDate = normalizeDateInput_(payload.date || payload.stay_date);
  const bookingRows = getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const historyIndex = buildGuestHistoryIndex_(bookingRows);
  const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));

  const matches = bookingRows
    .filter(function(row) {
      if (bookingId && String(row.booking_id || '').trim().toLowerCase().indexOf(bookingId) === -1) return false;
      if (guestName && String(row.guest_name || '').trim().toLowerCase().indexOf(guestName) === -1) return false;
      if (String(payload.source || '').trim() && normalizeBookingSource_(row.source || '') !== source) return false;
      if (targetDate) {
        const checkIn = normalizeDateInput_(row.check_in);
        const checkOut = normalizeDateInput_(row.check_out);
        if (!checkIn || !checkOut) return false;
        const target = stripTime_(targetDate).getTime();
        if (target < stripTime_(checkIn).getTime() || target >= stripTime_(checkOut).getTime()) return false;
      }
      return true;
    })
    .sort(function(a, b) {
      const aDate = normalizeDateInput_(a.check_in);
      const bDate = normalizeDateInput_(b.check_in);
      return (bDate ? bDate.getTime() : 0) - (aDate ? aDate.getTime() : 0);
    });
  const results = matches
    .slice(0, 50)
    .map(function(row) {
      return mapBookingRowToAdminDisplay_(row, {
        historyIndex: historyIndex,
        roomIndex: roomIndex
      });
    });

  return {
    ok: true,
    count: matches.length,
    hasMore: matches.length > 50,
    results: results
  };
}

function adminRefreshReporting() {
  return refreshMiniPmsReporting();
}

function adminRefreshAvailabilityCache() {
  return refreshAvailabilityCache(180);
}

function adminRebuildDailyStats() {
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  const dailyStats = rebuildDailyStats_(spreadsheet);
  const snapshot = rebuildOtbSnapshot_(spreadsheet);
  const dashboard = rebuildDashboard_(spreadsheet);

  return {
    ok: true,
    dailyStats: dailyStats,
    snapshot: snapshot,
    dashboard: dashboard
  };
}

function getFrontDeskDashboardData(selectedDate) {
  const frontDeskStartedAt = Date.now();
  logTiming_('frontDesk:start');
  try {
  return withRequestReadCache_(function() {
    const payload = selectedDate && typeof selectedDate === 'object' && !Array.isArray(selectedDate)
      ? selectedDate
      : {};
    const effectiveDate = selectedDate && typeof selectedDate === 'object' && !Array.isArray(selectedDate)
      ? (selectedDate.selected_date || selectedDate.selectedDate || selectedDate.date || '')
      : selectedDate;
    const uiMode = String(payload.ui_mode || payload.uiMode || '').trim().toLowerCase();
    const includeHeavyExtras = toBoolean_(payload.include_heavy_extras || payload.includeHeavyExtras || payload.full_refresh || payload.fullRefresh);
    if (uiMode === 'staff') {
      const roomRows = getRoomsMasterRows_({ activeOnly: false });
      const staffDate = normalizeFrontDeskDate_(effectiveDate);
      return buildFrontDeskDashboardData_(effectiveDate, {
        staffMode: true,
        bookings: getFrontDeskBookingRowsForDate_(staffDate),
        bookingsAreOperationalSubset: true,
        roomIndex: buildRoomMasterIndex_(roomRows),
        defaultBookingCurrency: DEFAULT_CURRENCY,
        reportingCurrency: getReportingCurrency_(),
        fastOperational: true,
        skipGuestHistory: true
      });
    }
    if (!includeHeavyExtras) {
      const roomRows = getRoomsMasterRows_({ activeOnly: false });
      const frontDeskDate = normalizeFrontDeskDate_(effectiveDate);
      const operationalBookings = getFrontDeskBookingRowsForDate_(frontDeskDate);
      return buildFrontDeskDashboardData_(effectiveDate, {
        staffMode: false,
        bookings: operationalBookings,
        bookingsAreOperationalSubset: true,
        roomIndex: buildRoomMasterIndex_(roomRows),
        defaultBookingCurrency: getDefaultBookingCurrency_(),
        reportingCurrency: getReportingCurrency_(),
        fastOperational: true,
        skipGuestHistory: true,
        skipAvailabilitySnapshot: true
      });
    }
    return buildFrontDeskDashboardData_(effectiveDate, {
      staffMode: false,
      fastOperational: true,
      skipGuestHistory: false
    });
  });
  } finally {
    logTiming_('frontDesk:return', frontDeskStartedAt);
  }
}

function getArrivalsForDate(selectedDate) {
  return withRequestReadCache_(function() {
    const date = normalizeFrontDeskDate_(selectedDate);
    const results = buildFrontDeskBookingLists_(date).arrivals;
    return {
      ok: true,
      selectedDate: formatDateKey_(date),
      count: results.length,
      results: results
    };
  });
}

function getArrivalsPlanningData(input) {
  return withRequestReadCache_(function() {
    const payload = input || {};
    const anchorDate = normalizeFrontDeskDate_(payload.start_date || payload.startDate || payload.selectedDate || payload.selected_date);
    const rangeKey = String(payload.range || payload.arrivals_range || payload.rangeKey || 'today').trim().toLowerCase();

    if (rangeKey === 'today') {
      const todayResults = buildFrontDeskBookingLists_(anchorDate).arrivals;
      return {
        ok: true,
        selectedDate: formatDateKey_(anchorDate),
        rangeKey: 'today',
        rangeLabel: 'Today',
        count: todayResults.length,
        results: todayResults
      };
    }

    const endExclusive = rangeKey === 'future'
      ? null
      : addDays_(anchorDate, rangeKey === '7' ? 7 : rangeKey === '14' ? 14 : rangeKey === '30' ? 30 : 7);
    const bookings = getSheetObjects_(SHEET_NAMES.BOOKINGS);
    const defaultBookingCurrency = getDefaultBookingCurrency_();
    const reportingCurrency = getReportingCurrency_();
    const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
    const filteredBookings = bookings.filter(function(row) {
        if (normalizeBookingStatus_(row.status) !== BOOKING_STATUS_CONFIRMED) return false;
        const checkIn = normalizeDateInput_(row.check_in);
        if (!checkIn) return false;
        const checkInTime = stripTime_(checkIn).getTime();
        if (checkInTime < stripTime_(anchorDate).getTime()) return false;
        if (endExclusive && checkInTime >= stripTime_(endExclusive).getTime()) return false;
        return true;
      });
    const historyIndex = buildGuestHistoryIndex_(bookings, {
      defaultBookingCurrency: defaultBookingCurrency,
      targetBookingIds: filteredBookings.map(function(row) {
        return String(row.booking_id || '').trim();
      }).filter(Boolean)
    });
    const results = filteredBookings
      .map(function(row) {
        return mapBookingRowToAdminDisplay_(row, {
          historyIndex: historyIndex,
          roomIndex: roomIndex,
          reportingCurrency: reportingCurrency,
          defaultBookingCurrency: defaultBookingCurrency
        });
      })
      .sort(function(a, b) {
        return compareArrivalDisplayRows_(a, b);
      });

    return {
      ok: true,
      selectedDate: formatDateKey_(anchorDate),
      rangeKey: rangeKey,
      rangeLabel: rangeKey === '7'
        ? 'Next 7 Days'
        : rangeKey === '14'
          ? 'Next 14 Days'
          : rangeKey === '30'
            ? 'Next 30 Days'
            : 'All Future Arrivals',
      count: results.length,
      results: results
    };
  });
}

function getDeparturesForDate(selectedDate) {
  return withRequestReadCache_(function() {
    const effectiveDate = selectedDate && typeof selectedDate === 'object' && !Array.isArray(selectedDate)
      ? (selectedDate.selected_date || selectedDate.selectedDate || selectedDate.date || '')
      : selectedDate;
    const date = normalizeFrontDeskDate_(effectiveDate);
    const results = buildFrontDeskBookingLists_(date).departures;
    return {
      ok: true,
      selectedDate: formatDateKey_(date),
      count: results.length,
      results: results
    };
  });
}

function getInHouseForDate(selectedDate) {
  return withRequestReadCache_(function() {
    const effectiveDate = selectedDate && typeof selectedDate === 'object' && !Array.isArray(selectedDate)
      ? (selectedDate.selected_date || selectedDate.selectedDate || selectedDate.date || '')
      : selectedDate;
    const date = normalizeFrontDeskDate_(effectiveDate);
    const results = buildFrontDeskBookingLists_(date).inHouse;
    return {
      ok: true,
      selectedDate: formatDateKey_(date),
      count: results.length,
      results: results
    };
  });
}

function getAvailabilityForDate(selectedDate, roomTypeId) {
  return buildAvailabilitySnapshotForDate_(selectedDate, roomTypeId);
}

function getAvailabilityCheckerData(input) {
  const payload = input || {};
  const checkIn = normalizeDateInput_(payload.check_in || payload.checkIn);
  const checkOut = normalizeDateInput_(payload.check_out || payload.checkOut);
  const roomTypeRaw = String(payload.room_type || payload.roomType || payload.room_type_id || payload.roomTypeId || '').trim();

  if (!checkIn) throw new Error('Check-in date is required.');
  if (!checkOut) throw new Error('Check-out date is required.');
  if (checkOut.getTime() <= checkIn.getTime()) throw new Error('Check-out must be after check-in.');

  if (roomTypeRaw) {
    const single = adminCheckAvailability({
      check_in: checkIn,
      check_out: checkOut,
      room_type: roomTypeRaw
    });
    single.roomTypeBreakdown = [{
      roomTypeId: single.roomTypeId,
      roomTypeName: single.roomTypeName,
      totalRooms: single.totalRooms,
      soldRooms: single.soldRooms,
      blockedRooms: single.blockedRooms,
      availableRooms: single.availableRooms,
      status: single.available ? 'Available' : 'Sold Out'
    }];
    return single;
  }

  return buildAvailabilityRangeOverview_(checkIn, checkOut);
}

function getAvailabilityPlannerData(input) {
  const plannerStartedAt = Date.now();
  logTiming_('planner:start');
  try {
  return withRequestReadCache_(function() {
    const payload = input || {};
    const startDate = normalizeDateInput_(payload.start_date || payload.startDate || payload.check_in || payload.checkIn || new Date());
    const rangeDays = normalizeAvailabilityPlannerDays_(payload.range_days || payload.rangeDays || payload.days);
    const endDate = addDays_(startDate, rangeDays);
    const overview = buildAvailabilityRangeOverview_(startDate, endDate, {
      includeInternalRows: true
    });

    return buildAvailabilityPlannerPayload_(startDate, rangeDays, overview, {
      bookingRows: overview.internalBookingRows || null
    });
  });
  } finally {
    logTiming_('planner:end', plannerStartedAt);
  }
}

function getReportsDashboardData(input) {
  const reportsStartedAt = Date.now();
  logTiming_('reports:start');
  try {
  const payload = (input && typeof input === 'object' && Object.prototype.toString.call(input) !== '[object Date]') ? input : {};
  if (toBoolean_(payload.allow_rebuild || payload.allowRebuild) && toBoolean_(payload.force_refresh || payload.forceRefresh)) {
    refreshMiniPmsReporting();
  }
  return buildReportsDashboardData_(input);
  } finally {
    logTiming_('reports:return', reportsStartedAt);
  }
}

function setupMiniPmsV1() {
  const spreadsheet = getSpreadsheet_();
  const ensured = ensureMiniPmsStructure_(spreadsheet);
  return {
    ok: true,
    mode: 'structure-only',
    ensuredSheets: ensured,
    nextStep: 'Run refreshMiniPmsReporting() manually after setup completes.'
  };
}

function refreshMiniPmsReporting() {
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);

  const availabilityCache = refreshAvailabilityCache(OTB_SNAPSHOT_DAYS_AHEAD);
  const bookingNights = rebuildBookingNights_(spreadsheet);
  const dailyStats = rebuildDailyStats_(spreadsheet);
  const snapshot = rebuildOtbSnapshot_(spreadsheet);
  const dashboard = rebuildDashboard_(spreadsheet);

  return {
    ok: true,
    availabilityCache: availabilityCache,
    bookingNights: bookingNights,
    dailyStats: dailyStats,
    snapshot: snapshot,
    dashboard: dashboard
  };
}

function refreshMiniPmsDashboard() {
  const spreadsheet = getSpreadsheet_();
  assertMiniPmsReady_(spreadsheet, [
    SHEET_NAMES.BOOKINGS,
    SHEET_NAMES.REQUESTS,
    SHEET_NAMES.BOOKING_NIGHTS,
    SHEET_NAMES.DAILY_STATS,
    SHEET_NAMES.OTB_SNAPSHOTS,
    SHEET_NAMES.SETTINGS
  ]);
  return rebuildDashboard_(spreadsheet);
}

function snapshotOnTheBooks() {
  const spreadsheet = getSpreadsheet_();
  ensureMiniPmsStructure_(spreadsheet);
  return rebuildOtbSnapshot_(spreadsheet);
}

function setupSingleTimeBasedTrigger_(handlerFunction, schedule) {
  const handler = String(handlerFunction || '').trim();
  if (!handler) throw new Error('Trigger handler function is required.');
  const config = schedule || {};
  let deleted = 0;

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger && trigger.getHandlerFunction && trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
      deleted += 1;
    }
  });

  let builder = ScriptApp.newTrigger(handler).timeBased();
  if (config.everyMinutes) {
    builder = builder.everyMinutes(Math.max(1, Number(config.everyMinutes || 1)));
  } else if (config.everyHours) {
    builder = builder.everyHours(Math.max(1, Number(config.everyHours || 1)));
  } else {
    builder = builder.everyDays(1).atHour(Math.max(0, Math.min(23, Number(config.atHour || 6))));
  }
  const trigger = builder.create();
  const scheduleLabel = config.everyMinutes
    ? ('every ' + config.everyMinutes + ' minute(s)')
    : (config.everyHours ? ('every ' + config.everyHours + ' hour(s)') : ('daily at hour ' + Math.max(0, Math.min(23, Number(config.atHour || 6)))));

  return {
    ok: true,
    handler: handler,
    deletedExisting: deleted,
    triggerId: trigger && trigger.getUniqueId ? trigger.getUniqueId() : '',
    schedule: scheduleLabel
  };
}

function setupAvailabilityCacheRefreshTrigger() {
  return setupSingleTimeBasedTrigger_('refreshAvailabilityCacheScheduled', { everyHours: 6 });
}

function setupWebsiteEmailProcessorTrigger() {
  return setupSingleTimeBasedTrigger_(WEBSITE_EMAIL_FOLLOW_UP_HANDLER, { everyMinutes: 5 });
}

function getWebsiteEmailProcessorTriggerStatus_() {
  let triggers = [];
  try {
    triggers = ScriptApp.getProjectTriggers();
  } catch (error) {
    return {
      ok: false,
      exists: false,
      count: 0,
      handler: WEBSITE_EMAIL_FOLLOW_UP_HANDLER,
      verificationUnavailable: true,
      message: 'Trigger verification requires owner/admin authorization in Apps Script editor.'
    };
  }
  const matches = triggers.filter(function(trigger) {
    return trigger && trigger.getHandlerFunction && trigger.getHandlerFunction() === WEBSITE_EMAIL_FOLLOW_UP_HANDLER;
  });
  return {
    ok: matches.length === 1,
    exists: matches.length > 0,
    count: matches.length,
    handler: WEBSITE_EMAIL_FOLLOW_UP_HANDLER,
    verificationUnavailable: false,
    message: matches.length === 1
      ? 'Website email processor trigger is installed.'
      : (matches.length ? 'Multiple website email processor triggers were found.' : 'Website email processor trigger was not found.')
  };
}

function adminVerifyWebsiteEmailProcessorTrigger(input) {
  return getWebsiteEmailProcessorTriggerStatus_();
}

function getPublicDeploymentSafetyChecklist() {
  return {
    ok: true,
    message: 'Public hosting must include only public website assets. Do not upload PMS/admin source files.',
    doNotPubliclyHost: [
      'Code.gs.gs',
      'AdminPanel.html',
      'PMS/admin source files',
      'Google Apps Script source exports',
      'spreadsheet IDs or Apps Script deployment credentials'
    ],
    allowedPublicFiles: [
      'index.html',
      'book.html',
      'public image assets'
    ]
  };
}

function setupDailyPmsSummaryTrigger() {
  return setupSingleTimeBasedTrigger_('sendDailyPmsSummaryEmail', { atHour: 8 });
}

function setupNightlyOtbSnapshotTrigger() {
  return setupSingleTimeBasedTrigger_('runNightlyOtbSnapshot', { atHour: 2 });
}

function setupPreGoLiveOperationalTriggers() {
  return {
    ok: true,
    availabilityCache: setupAvailabilityCacheRefreshTrigger(),
    websiteEmailProcessor: setupWebsiteEmailProcessorTrigger(),
    dailySummary: setupDailyPmsSummaryTrigger(),
    nightlyOtbSnapshot: setupNightlyOtbSnapshotTrigger()
  };
}

function adminSetupOperationalSafetyTriggers(input) {
  return setupPreGoLiveOperationalTriggers(input || {});
}

function getPmsAdminAlertEmails_() {
  const emails = [];
  try {
    const propertyValue = String(PropertiesService.getScriptProperties().getProperty(PMS_ADMIN_EMAILS_PROPERTY_KEY) || '').trim();
    propertyValue.split(/[;,]/).forEach(function(email) {
      const value = String(email || '').trim();
      if (value) emails.push(value);
    });
  } catch (error) {
    // Continue to Settings fallback.
  }

  try {
    const contact = getWebsiteContactDetails_();
    if (contact.notificationEmail) emails.push(contact.notificationEmail);
    if (contact.email) emails.push(contact.email);
  } catch (error) {
    // Alerting must not fail because Settings are temporarily unavailable.
  }

  const seen = {};
  return emails.filter(function(email) {
    const value = String(email || '').trim().toLowerCase();
    if (!value || seen[value] || !isValidEmailAddress_(value)) return false;
    seen[value] = true;
    return true;
  });
}

function sendSilentFailureAlert_(jobName, error, details) {
  const name = String(jobName || 'Scheduled PMS job').trim();
  const message = String(error && error.message || error || 'Unknown error');
  const recipients = getPmsAdminAlertEmails_();
  if (!recipients.length) {
    Logger.log('Silent failure alert not sent; no admin email configured. Job=' + name + ' Error=' + message);
    return { ok: false, skipped: true, reason: 'No admin email configured.' };
  }

  const cacheKey = 'silentFailureAlert:' + name.toLowerCase().replace(/[^a-z0-9]+/g, ':').slice(0, 80);
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get(cacheKey)) {
      Logger.log('Silent failure alert suppressed to avoid spam. Job=' + name + ' Error=' + message);
      return { ok: true, suppressed: true };
    }
    cache.put(cacheKey, 'sent', 6 * 60 * 60);
  } catch (cacheError) {
    // Continue without throttling if cache is unavailable.
  }

  const body = [
    'Roza PMS scheduled job failed.',
    '',
    'Job: ' + name,
    'Time: ' + Utilities.formatDate(new Date(), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm:ss'),
    'Error: ' + message,
    '',
    details ? ('Details: ' + String(details)) : ''
  ].filter(function(line) { return line !== ''; }).join('\n');

  try {
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: 'Roza PMS alert: ' + name + ' failed',
      body: body
    });
  } catch (mailError) {
    Logger.log('Silent failure alert email failed. Job=' + name + ' Error=' + String(mailError && mailError.message || mailError));
    return { ok: false, skipped: true, reason: 'Alert email failed.' };
  }

  return { ok: true, recipients: recipients };
}

function refreshAvailabilityCacheScheduled() {
  const startedAt = Date.now();
  logTiming_('availabilityCacheScheduled:start');
  try {
    const result = refreshAvailabilityCache(180);
    Logger.log('Availability cache scheduled refresh completed: ' + JSON.stringify(result));
    logTiming_('availabilityCacheScheduled:return', startedAt);
    return result;
  } catch (error) {
    sendSilentFailureAlert_('Availability cache refresh', error);
    throw error;
  }
}

function countWebsiteEmailFollowUpStates_(bookingRows) {
  const counts = {
    pending: 0,
    failed: 0,
    sending: 0,
    rows: []
  };

  (bookingRows || []).forEach(function(row) {
    if (normalizeBookingSource_(row.source || '') !== 'Direct Website') return;
    const bookingStatus = normalizeBookingStatus_(row.status);
    if (bookingStatus === BOOKING_STATUS_CANCELLED || bookingStatus === BOOKING_STATUS_NO_SHOW) return;
    const state = getWebsiteEmailFollowUpStateFromNotes_(row.internal_notes);
    if (state !== 'pending' && state !== 'failed' && state !== 'sending') return;
    counts[state] += 1;
    counts.rows.push({
      bookingId: String(row.booking_id || '').trim(),
      guestName: String(row.guest_name || '').trim(),
      state: state
    });
  });

  return counts;
}

function bookingSummaryLine_(row) {
  const bookingId = String(row.booking_id || '').trim();
  const guestName = String(row.guest_name || 'Guest').trim();
  const room = String(row.room_identifier || row.room_type_name || row.room_type_id || 'No room').trim();
  const balance = calculateBalanceDue_(getOperationalBookingValueGbp_(row), roundCurrency_(row.amount_paid || 0));
  return guestName + ' | ' + room + (bookingId ? ' | ' + bookingId : '') + (balance > 0.009 ? ' | due ' + formatMoneyValue_(balance, getReportingCurrency_()) : '');
}

function listSummaryLines_(rows, limit) {
  const safeRows = rows || [];
  const max = Math.max(1, Number(limit || 12));
  if (!safeRows.length) return ['None'];
  const lines = safeRows.slice(0, max).map(bookingSummaryLine_);
  if (safeRows.length > max) lines.push('+' + (safeRows.length - max) + ' more');
  return lines;
}

function buildDailyPmsSummaryData_() {
  const today = stripTime_(new Date());
  const todayKey = formatDateKey_(today);
  const bookings = getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const arrivals = bookings.filter(function(row) {
    const checkIn = normalizeDateInput_(row.check_in);
    return normalizeBookingStatus_(row.status) === BOOKING_STATUS_CONFIRMED &&
      checkIn && formatDateKey_(checkIn) === todayKey;
  });
  const departures = bookings.filter(function(row) {
    const checkOut = normalizeDateInput_(row.check_out);
    return normalizeBookingStatus_(row.status) === BOOKING_STATUS_IN_HOUSE &&
      checkOut && formatDateKey_(checkOut) === todayKey;
  });
  const inHouse = bookings.filter(function(row) {
    const status = normalizeBookingStatus_(row.status);
    if (status !== BOOKING_STATUS_IN_HOUSE) return false;
    const checkIn = normalizeDateInput_(row.check_in);
    const checkOut = normalizeDateInput_(row.check_out);
    return checkIn && checkOut && stripTime_(checkIn).getTime() <= today.getTime() && today.getTime() < stripTime_(checkOut).getTime();
  });
  const unpaid = bookings.filter(function(row) {
    const status = normalizeBookingStatus_(row.status);
    if (status === BOOKING_STATUS_CANCELLED || status === BOOKING_STATUS_NO_SHOW || status === BOOKING_STATUS_CHECKED_OUT) return false;
    return calculateBalanceDue_(getOperationalBookingValueGbp_(row), roundCurrency_(row.amount_paid || 0)) > 0.009;
  });
  const availability = buildAvailabilitySnapshotForDate_(today, '', {
    bookings: bookings
  });
  let otbToday = 0;
  try {
    getSheetObjects_(SHEET_NAMES.DAILY_STATS).forEach(function(row) {
      const stayDate = normalizeDateInput_(row.stay_date);
      if (stayDate && formatDateKey_(stayDate) === todayKey) {
        otbToday += Number(row.room_revenue || 0);
      }
    });
  } catch (error) {
    otbToday = 0;
  }
  const emailStates = countWebsiteEmailFollowUpStates_(bookings);
  const health = runPmsHealthCheck({ skipAvailabilityDeepCheck: true });

  return {
    date: todayKey,
    arrivals: arrivals,
    departures: departures,
    inHouse: inHouse,
    unpaid: unpaid,
    availability: availability,
    otbToday: roundCurrency_(otbToday),
    emailStates: emailStates,
    health: health
  };
}

function buildDailyPmsSummaryEmailBody_(summary) {
  const availability = summary.availability && summary.availability.totals ? summary.availability.totals : {};
  const emailStates = summary.emailStates || {};
  const health = summary.health || { warnings: [], errors: [] };

  return [
    'Roza PMS Daily Summary - ' + summary.date,
    '',
    'Arrivals today: ' + summary.arrivals.length,
    listSummaryLines_(summary.arrivals, 12).map(function(line) { return '  - ' + line; }).join('\n'),
    '',
    'Departures today: ' + summary.departures.length,
    listSummaryLines_(summary.departures, 12).map(function(line) { return '  - ' + line; }).join('\n'),
    '',
    'In-house: ' + summary.inHouse.length,
    'Unpaid active balances: ' + summary.unpaid.length,
    listSummaryLines_(summary.unpaid, 12).map(function(line) { return '  - ' + line; }).join('\n'),
    '',
    'Rooms available today: ' + Number(availability.availableRooms || 0) + ' of ' + Number(availability.totalRooms || 0),
    'OTB revenue today: ' + formatMoneyValue_(summary.otbToday || 0, getReportingCurrency_()),
    'Website email follow-up: pending ' + Number(emailStates.pending || 0) + ', failed ' + Number(emailStates.failed || 0) + ', sending ' + Number(emailStates.sending || 0),
    '',
    'Health check: ' + (health.ok ? 'OK' : 'Needs attention'),
    'Errors: ' + (health.errors && health.errors.length ? health.errors.join(' | ') : 'None'),
    'Warnings: ' + (health.warnings && health.warnings.length ? health.warnings.join(' | ') : 'None')
  ].join('\n');
}

function sendDailyPmsSummaryEmail(input) {
  const startedAt = Date.now();
  logTiming_('dailySummary:start');
  try {
    const recipients = getPmsAdminAlertEmails_();
    if (!recipients.length) {
      Logger.log('Daily PMS summary skipped; no admin email configured.');
      return { ok: false, skipped: true, reason: 'No admin email configured.' };
    }
    const summary = buildDailyPmsSummaryData_();
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: 'Roza PMS Daily Summary - ' + summary.date,
      body: buildDailyPmsSummaryEmailBody_(summary)
    });
    logTiming_('dailySummary:return', startedAt);
    return {
      ok: true,
      date: summary.date,
      recipients: recipients,
      arrivals: summary.arrivals.length,
      departures: summary.departures.length,
      inHouse: summary.inHouse.length,
      unpaid: summary.unpaid.length,
      healthOk: summary.health.ok
    };
  } catch (error) {
    sendSilentFailureAlert_('Daily PMS summary email', error);
    throw error;
  }
}

function adminSendDailyPmsSummaryEmail(input) {
  return sendDailyPmsSummaryEmail(input || {});
}

function runPmsHealthCheck(input) {
  const startedAt = Date.now();
  logTiming_('healthCheck:start');
  const warnings = [];
  const errors = [];
  const today = stripTime_(new Date());
  const todayKey = formatDateKey_(today);

  try {
    const bookings = getSheetObjects_(SHEET_NAMES.BOOKINGS);
    const seenBookingIds = {};
    bookings.forEach(function(row, index) {
      const rowNumber = index + 2;
      const bookingId = String(row.booking_id || '').trim();
      if (!bookingId) {
        errors.push('Booking row ' + rowNumber + ' is missing booking_id. Manual sheet cleanup is required; this health check does not modify booking data.');
      } else if (seenBookingIds[bookingId]) {
        errors.push('Duplicate booking_id ' + bookingId + ' on rows ' + seenBookingIds[bookingId] + ' and ' + rowNumber + '.');
      } else {
        seenBookingIds[bookingId] = rowNumber;
      }

      const status = normalizeBookingStatus_(row.status);
      const checkIn = normalizeDateInput_(row.check_in);
      const checkOut = normalizeDateInput_(row.check_out);
      const activeFutureOrCurrent = checkOut && stripTime_(checkOut).getTime() >= today.getTime();
      if ((status === BOOKING_STATUS_CONFIRMED || status === BOOKING_STATUS_IN_HOUSE) && activeFutureOrCurrent && !String(row.room_type_id || row.room_type_name || '').trim()) {
        errors.push('Booking ' + (bookingId || ('row ' + rowNumber)) + ' is active/current but missing room type.');
      }
      if (status === BOOKING_STATUS_CONFIRMED && checkIn && stripTime_(checkIn).getTime() >= today.getTime() && !String(row.room_identifier || '').trim()) {
        warnings.push('Arrival ' + (bookingId || ('row ' + rowNumber)) + ' has no room assignment.');
      }
      if (status === BOOKING_STATUS_IN_HOUSE && !String(row.room_identifier || '').trim()) {
        errors.push('Checked-in booking ' + (bookingId || ('row ' + rowNumber)) + ' has no room assignment. Manual correction is required.');
      }
      if (status === BOOKING_STATUS_IN_HOUSE && !hasRecordedPaymentMethod_(row.payment_method)) {
        errors.push('Checked-in booking ' + (bookingId || ('row ' + rowNumber)) + ' has no method of payment. Manual correction is required.');
      }
    });

    const emailStates = countWebsiteEmailFollowUpStates_(bookings);
    if (emailStates.failed) warnings.push(String(emailStates.failed) + ' website booking email follow-up(s) failed.');
    if (emailStates.pending) warnings.push(String(emailStates.pending) + ' website booking email follow-up(s) pending.');
    if (emailStates.sending) warnings.push(String(emailStates.sending) + ' website booking email follow-up(s) currently sending.');

    const websiteEmailTrigger = getWebsiteEmailProcessorTriggerStatus_();
    if (websiteEmailTrigger.verificationUnavailable) {
      warnings.push(websiteEmailTrigger.message || 'Trigger verification requires owner/admin authorization in Apps Script editor.');
    } else if (!websiteEmailTrigger.exists) {
      errors.push('Website email processor trigger is missing. Run setupWebsiteEmailProcessorTrigger or adminSetupOperationalSafetyTriggers.');
    } else if (websiteEmailTrigger.count > 1) {
      warnings.push('Website email processor has ' + websiteEmailTrigger.count + ' triggers; run setupWebsiteEmailProcessorTrigger to reset to one.');
    }

    const activeRoomTypes = getActiveRoomTypeCatalog_();
    const baseRateRows = getSheetObjects_(SHEET_NAMES.BASE_RATES);
    const baseRatesByRoomType = {};
    baseRateRows.forEach(function(row) {
      const roomTypeId = String(row.room_type_id || '').trim();
      if (roomTypeId) baseRatesByRoomType[roomTypeId] = row;
    });
    activeRoomTypes.forEach(function(roomType) {
      const base = baseRatesByRoomType[roomType.roomTypeId];
      if (!base || Number(base.base_rate || 0) <= 0) {
        warnings.push('Missing active base rate for ' + (roomType.roomTypeName || roomType.roomTypeId) + '.');
      }
    });

    const cacheRows = getSheetObjects_(SHEET_NAMES.AVAILABILITY_CACHE);
    const todayCacheRoomTypes = {};
    let latestCacheDate = null;
    cacheRows.forEach(function(row) {
      const cacheDate = normalizeDateInput_(row.date);
      if (!cacheDate) return;
      if (!latestCacheDate || cacheDate.getTime() > latestCacheDate.getTime()) latestCacheDate = cacheDate;
      if (formatDateKey_(cacheDate) === todayKey) {
        todayCacheRoomTypes[String(row.room_type_id || '').trim()] = true;
      }
    });
    if (!cacheRows.length) {
      warnings.push('Availability cache has no rows.');
    } else if (latestCacheDate && stripTime_(latestCacheDate).getTime() < today.getTime()) {
      warnings.push('Availability cache is stale; latest date is ' + formatDateKey_(latestCacheDate) + '.');
    }
    activeRoomTypes.forEach(function(roomType) {
      if (!todayCacheRoomTypes[roomType.roomTypeId]) {
        warnings.push('Availability cache missing today for ' + (roomType.roomTypeName || roomType.roomTypeId) + '.');
      }
    });

    getSheetObjects_(SHEET_NAMES.COMMERCIAL_CONTROLS).forEach(function(row) {
      const active = String(row.active || '').trim().toLowerCase();
      if (active && active !== 'yes' && active !== 'true' && active !== 'active') return;
      const startDate = normalizeDateInput_(row.start_date);
      const endDate = normalizeDateInput_(row.end_date);
      const controlId = String(row.control_id || row.room_type_id || 'Commercial control').trim();
      if (!startDate || !endDate) {
        warnings.push(controlId + ' has an invalid commercial control date.');
        return;
      }
      if (endDate.getTime() < startDate.getTime()) {
        errors.push(controlId + ' has end_date before start_date.');
      }
    });
  } catch (error) {
    errors.push(String(error && error.message || error));
  }

  const result = {
    ok: errors.length === 0 && warnings.length === 0,
    checkedAt: Utilities.formatDate(new Date(), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm:ss'),
    warnings: warnings,
    errors: errors
  };
  logTiming_('healthCheck:return', startedAt);
  return result;
}

function adminRunPmsHealthCheck(input) {
  return runPmsHealthCheck(input || {});
}

function runNightlyOtbSnapshot() {
  const startedAt = Date.now();
  logTiming_('nightlyOtb:start');
  try {
    const spreadsheet = getSpreadsheet_();
    ensureMiniPmsStructure_(spreadsheet);
    const bookingNights = rebuildBookingNights_(spreadsheet);
    const dailyStats = rebuildDailyStats_(spreadsheet);
    const snapshot = rebuildOtbSnapshot_(spreadsheet);
    const result = {
      ok: true,
      bookingNights: bookingNights,
      dailyStats: dailyStats,
      snapshot: snapshot
    };
    Logger.log('Nightly OTB snapshot completed: ' + JSON.stringify(result));
    logTiming_('nightlyOtb:return', startedAt);
    return result;
  } catch (error) {
    sendSilentFailureAlert_('Nightly OTB snapshot', error);
    throw error;
  }
}

function buildBookingRefreshWindow_(validated) {
  return {
    startDate: validated.checkIn,
    endDate: validated.checkOut,
    roomTypeIds: [validated.roomTypeId]
  };
}

function buildBookingRecoveryNote_(errors) {
  return 'Recovery needed after booking creation: ' + (errors || []).join(' | ');
}

function buildDeferredBookingPostCommitResult_(config) {
  const result = {
    requestSync: {
      status: config.requestId && !config.skipRequestSync ? 'deferred' : 'skipped',
      error: ''
    },
    availabilityRefresh: {
      status: 'deferred',
      error: ''
    },
    reportingRefresh: {
      status: config.skipRefresh ? 'skipped' : 'deferred',
      error: ''
    },
    recoveryNeeded: true,
    recoveryMessage: '',
    internalNotes: String(config.internalNotes || '').trim()
  };
  const deferred = ['Availability cache refresh deferred'];
  if (config.requestId && !config.skipRequestSync) {
    deferred.push('Request sync deferred');
  }
  if (!config.skipRefresh) {
    deferred.push('Reporting refresh deferred');
  }
  result.recoveryMessage = 'Booking ' + config.bookingId + ' was created. Internal follow-up was deferred so guest confirmation could return quickly. ' + deferred.join(' | ');
  return result;
}

function runBookingPostCommitTasks_(config) {
  const result = {
    requestSync: {
      status: config.requestId && !config.skipRequestSync ? 'pending' : 'skipped',
      error: ''
    },
    availabilityRefresh: {
      status: config.skipAvailabilityRefresh ? 'skipped' : 'pending',
      error: ''
    },
    reportingRefresh: {
      status: config.skipRefresh ? 'skipped' : 'pending',
      error: ''
    },
    recoveryNeeded: false,
    recoveryMessage: '',
    internalNotes: String(config.internalNotes || '').trim()
  };
  const unresolvedErrors = [];

  if (config.requestId && !config.skipRequestSync) {
    try {
      syncRequestAsConverted_(config.requestId, {
        bookingId: config.bookingId,
        convertedAt: config.createdAt,
        source: config.source,
        bookingValue: config.bookingValue
      });
      result.requestSync.status = 'ok';
    } catch (error) {
      result.requestSync.status = 'failed';
      result.requestSync.error = String(error && error.message || error || 'Request sync failed.');
      unresolvedErrors.push('Request sync failed: ' + result.requestSync.error);
    }
  }

  if (!config.skipAvailabilityRefresh) {
    try {
      refreshAvailabilityCacheWindow_(buildBookingRefreshWindow_(config.validated));
      result.availabilityRefresh.status = 'ok';
    } catch (error) {
      result.availabilityRefresh.status = 'failed';
      result.availabilityRefresh.error = String(error && error.message || error || 'Availability refresh failed.');
    }
  }

  if (!config.skipRefresh) {
    try {
      refreshMiniPmsReporting();
      result.reportingRefresh.status = 'ok';
      if (result.availabilityRefresh.status === 'failed') {
        result.availabilityRefresh.status = 'superseded';
        result.availabilityRefresh.error = '';
      }
    } catch (error) {
      result.reportingRefresh.status = 'failed';
      result.reportingRefresh.error = String(error && error.message || error || 'Reporting refresh failed.');
      unresolvedErrors.push('Reporting refresh failed: ' + result.reportingRefresh.error);
      if (result.availabilityRefresh.status === 'failed') {
        unresolvedErrors.unshift('Availability refresh failed: ' + result.availabilityRefresh.error);
      }
    }
  } else if (result.availabilityRefresh.status === 'failed') {
    unresolvedErrors.push('Availability refresh failed: ' + result.availabilityRefresh.error);
  }

  if (unresolvedErrors.length) {
    result.recoveryNeeded = true;
    result.recoveryMessage = 'Booking ' + config.bookingId + ' was created, but internal follow-up is still needed. ' + unresolvedErrors.join(' | ');
    try {
      result.internalNotes = mergeOperationalNotes_(result.internalNotes, buildBookingRecoveryNote_(unresolvedErrors));
      updateObjectRow_(config.sheet, config.rowNumber, {
        internal_notes: result.internalNotes
      });
    } catch (noteError) {
      const noteMessage = String(noteError && noteError.message || noteError || 'Recovery note could not be saved.');
      result.recoveryMessage += ' Recovery note could not be saved automatically: ' + noteMessage;
    }
  }

  return result;
}

function buildRoomTypeFilterMap_(roomTypeIds) {
  const map = {};
  (roomTypeIds || []).forEach(function(roomTypeId) {
    const id = String(roomTypeId || '').trim();
    if (id) map[id] = true;
  });
  return map;
}

function getSheetRowsForHeadersContiguous_(sheet, headerMap, headerKeys) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const keys = [];
  (headerKeys || []).forEach(function(headerKey) {
    const key = String(headerKey || '').trim();
    if (key && Object.prototype.hasOwnProperty.call(headerMap, key) && keys.indexOf(key) === -1) {
      keys.push(key);
    }
  });
  if (!keys.length) return [];

  const indexes = keys.map(function(key) {
    return headerMap[key];
  });
  const startIndex = Math.min.apply(null, indexes);
  const endIndex = Math.max.apply(null, indexes);
  const values = sheet.getRange(2, startIndex + 1, lastRow - 1, endIndex - startIndex + 1).getValues();
  return values.map(function(row) {
    const object = {};
    keys.forEach(function(key) {
      object[key] = row[headerMap[key] - startIndex];
    });
    return object;
  });
}

function getBookingRowsForAvailabilityWindow_(bookingsSheet, checkIn, checkOut, roomTypeIds) {
  const startDate = normalizeDateInput_(checkIn);
  const endDate = normalizeDateInput_(checkOut);
  if (!startDate || !endDate) return [];
  const targetRoomTypes = buildRoomTypeFilterMap_(roomTypeIds);
  const hasRoomTypeFilter = Object.keys(targetRoomTypes).length > 0;
  const headerMap = getHeaderMap_(bookingsSheet);
  return getSheetRowsForHeadersContiguous_(bookingsSheet, headerMap, [
    'booking_id',
    'guest_name',
    'check_in',
    'check_out',
    'actual_check_out_date',
    'room_type_id',
    'room_identifier',
    'bed_setup',
    'adults',
    'children',
    'guests',
    'qty_rooms',
    'status'
  ]).filter(function(row) {
    const roomTypeId = String(row.room_type_id || '').trim();
    if (!roomTypeId || (hasRoomTypeFilter && !targetRoomTypes[roomTypeId])) return false;
    if (!isRevenueCountedStatus_(normalizeBookingStatus_(row.status))) return false;
    const rowCheckIn = normalizeDateInput_(row.check_in);
    const rowCheckOut = getInventoryEffectiveCheckOut_(row);
    return staysOverlap_(startDate, endDate, rowCheckIn, rowCheckOut);
  });
}

function buildDemandUnitMapForAvailabilityWindow_(bookingRows, checkIn, checkOut, roomTypeIds, roomIndex) {
  const startDate = normalizeDateInput_(checkIn);
  const endDate = normalizeDateInput_(checkOut);
  if (!startDate || !endDate) return {};

  const targetRoomTypes = buildRoomTypeFilterMap_(roomTypeIds);
  const hasRoomTypeFilter = Object.keys(targetRoomTypes).length > 0;
  const effectiveRoomIndex = roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const map = {};

  (bookingRows || []).forEach(function(row) {
    const bookingId = String(row.booking_id || '').trim();
    const roomTypeId = String(row.room_type_id || '').trim();
    if (!roomTypeId || (hasRoomTypeFilter && !targetRoomTypes[roomTypeId])) return;
    if (BOOKING_STATUSES_COUNTED.indexOf(normalizeBookingStatus_(row.status)) === -1) return;

    const rowCheckIn = normalizeDateInput_(row.check_in);
    const rowCheckOut = getInventoryEffectiveCheckOut_(row);
    if (!rowCheckIn || !rowCheckOut || !staysOverlap_(startDate, endDate, rowCheckIn, rowCheckOut)) return;

    const qtyRooms = Math.max(1, Number(row.qty_rooms || 1));
    const totalGuests = Math.max(1, Number(row.guests || Number(row.adults || 0) + Number(row.children || 0) || 1));
    const guests = getGuestsPerRoom_(totalGuests, qtyRooms);
    const bedSetup = normalizeBedSetup_(row.bed_setup || '');
    const assignedRoom = qtyRooms === 1
      ? resolveAssignedRoomFromIndex_(effectiveRoomIndex, row.room_identifier, roomTypeId)
      : null;
    const overlapStart = rowCheckIn.getTime() > startDate.getTime() ? rowCheckIn : startDate;
    const overlapEnd = rowCheckOut.getTime() < endDate.getTime() ? rowCheckOut : endDate;

    for (let date = stripTime_(overlapStart); date.getTime() < stripTime_(overlapEnd).getTime(); date = addDays_(date, 1)) {
      const key = getDateRoomKey_(date, roomTypeId);
      if (!map[key]) map[key] = [];
      for (let index = 0; index < qtyRooms; index++) {
        map[key].push({
          unitId: bookingId + '#' + index,
          bookingId: bookingId,
          guests: guests,
          bedSetup: bedSetup,
          fixedRoomId: assignedRoom && index === 0 ? assignedRoom.roomId : ''
        });
      }
    }
  });

  return map;
}

function getBlockedRowsForAvailabilityWindow_(spreadsheet, checkIn, checkOut, roomTypeIds) {
  const startDate = normalizeDateInput_(checkIn);
  const endDate = normalizeDateInput_(checkOut);
  if (!startDate || !endDate) return [];
  const targetRoomTypes = buildRoomTypeFilterMap_(roomTypeIds);
  const hasRoomTypeFilter = Object.keys(targetRoomTypes).length > 0;
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BLOCKED_DATES);
  const headerMap = getHeaderMap_(sheet);
  return getSheetRowsForHeadersContiguous_(sheet, headerMap, [
    'date',
    'room_type_id',
    'qty_blocked',
    'status'
  ]).filter(function(row) {
    const roomTypeId = String(row.room_type_id || '').trim();
    if (!roomTypeId || (hasRoomTypeFilter && !targetRoomTypes[roomTypeId])) return false;
    const status = String(row.status || 'Active').trim().toLowerCase();
    if (status && status !== 'active' && status !== 'open') return false;
    const blockDate = normalizeDateInput_(row.date);
    return blockDate && blockDate.getTime() >= startDate.getTime() && blockDate.getTime() < endDate.getTime();
  });
}

function getBookingRowsForInventoryRecheck_(bookingsSheet, checkIn, checkOut, roomTypeId) {
  const startDate = normalizeDateInput_(checkIn);
  const endDate = normalizeDateInput_(checkOut);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const headerMap = getHeaderMap_(bookingsSheet);
  return getSheetRowsForHeaders_(bookingsSheet, headerMap, [
    'booking_id',
    'guest_name',
    'check_in',
    'check_out',
    'actual_check_out_date',
    'room_type_id',
    'room_identifier',
    'bed_setup',
    'adults',
    'children',
    'guests',
    'qty_rooms',
    'status'
  ]).filter(function(row) {
    if (String(row.room_type_id || '').trim() !== targetRoomTypeId) return false;
    if (BOOKING_STATUSES_COUNTED.indexOf(normalizeBookingStatus_(row.status)) === -1) return false;
    const rowCheckIn = normalizeDateInput_(row.check_in);
    const rowCheckOut = getInventoryEffectiveCheckOut_(row);
    return staysOverlap_(startDate, endDate, rowCheckIn, rowCheckOut);
  });
}

function getBlockedRowsForInventoryRecheck_(spreadsheet, checkIn, checkOut, roomTypeId) {
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BLOCKED_DATES);
  const startDate = normalizeDateInput_(checkIn);
  const endDate = normalizeDateInput_(checkOut);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const headerMap = getHeaderMap_(sheet);
  return getSheetRowsForHeaders_(sheet, headerMap, [
    'date',
    'room_type_id',
    'qty_blocked',
    'status'
  ]).filter(function(row) {
    if (String(row.room_type_id || '').trim() !== targetRoomTypeId) return false;
    const status = String(row.status || 'Active').trim().toLowerCase();
    if (status && status !== 'active' && status !== 'open') return false;
    const blockDate = normalizeDateInput_(row.date);
    return blockDate && blockDate.getTime() >= startDate.getTime() && blockDate.getTime() < endDate.getTime();
  });
}

function createManualBooking(input) {
  const createManualBookingStartedAt = Date.now();
  const validateManualStartedAt = Date.now();
  const validated = validateManualBookingInput_(input);
  const isWebsiteDirectBooking = normalizeBookingSource_(validated.source || '') === 'Direct Website';
  const logCreateManualBookingTiming = function(label, startedAt) {
    if (isWebsiteDirectBooking) logTiming_(label, startedAt);
  };
  logCreateManualBookingTiming('createManualBooking:validate', validateManualStartedAt);
  const spreadsheet = getSpreadsheet_();
  if (!validated.skipStructureEnsure) {
    assertMiniPmsReady_(spreadsheet, getAvailabilityOperationSheetNames_());
  }
  const bookingsSheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  const now = new Date();
  const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const lock = LockService.getScriptLock();

  let bookingId;
  let rowNumber;

  const lockWaitStartedAt = Date.now();
  logCreateManualBookingTiming('createManualBooking:lockStart');
  lock.waitLock(30000);
  logTiming_(isWebsiteDirectBooking ? 'websiteBooking:lockWait' : 'createBooking:lockWait', lockWaitStartedAt);
  logCreateManualBookingTiming('createManualBooking:lockWait', lockWaitStartedAt);
  try {
    const readBookingsStartedAt = Date.now();
    const bookingRows = validated.fastInventoryRows
      ? getBookingRowsForInventoryRecheck_(bookingsSheet, validated.checkIn, validated.checkOut, validated.roomTypeId)
      : getSheetObjects_(SHEET_NAMES.BOOKINGS);
    logCreateManualBookingTiming('createManualBooking:readBookings', readBookingsStartedAt);
    const readBlockedStartedAt = Date.now();
    const blockedRows = validated.fastInventoryRows
      ? getBlockedRowsForInventoryRecheck_(spreadsheet, validated.checkIn, validated.checkOut, validated.roomTypeId)
      : getSheetObjects_(SHEET_NAMES.BLOCKED_DATES);
    logCreateManualBookingTiming('createManualBooking:readBlocked', readBlockedStartedAt);
    const buildBlockedMapStartedAt = Date.now();
    const blockedCountMap = buildBlockedDateQtyMap_(blockedRows, {
      startDate: validated.checkIn,
      endDateExclusive: validated.checkOut
    }).byDateRoom;
    logCreateManualBookingTiming('createManualBooking:buildBlockedMap', buildBlockedMapStartedAt);
    const buildDemandMapStartedAt = Date.now();
    const demandUnitsByDateRoom = buildDemandUnitMapForAvailabilityWindow_(bookingRows, validated.checkIn, validated.checkOut, [validated.roomTypeId], roomIndex);
    logCreateManualBookingTiming('createManualBooking:buildDemandMap', buildDemandMapStartedAt);
    const finalInventoryStartedAt = Date.now();
    const liveAvailability = buildStayAvailabilityInventorySnapshot_(validated.checkIn, validated.checkOut, validated.roomTypeId, validated.guests, {
      qtyRooms: validated.qtyRooms,
      bedSetup: validated.bedSetup,
      sellableProductId: validated.sellableProductId,
      bookingRows: bookingRows,
      blockedRows: blockedRows,
      blockedCountMap: blockedCountMap,
      demandUnitsByDateRoom: demandUnitsByDateRoom,
      roomIndex: roomIndex
    });
    logCreateManualBookingTiming('createManualBooking:inventorySnapshot', finalInventoryStartedAt);
    logTiming_(isWebsiteDirectBooking ? 'websiteBooking:finalInventoryRecheck' : 'createBooking:finalInventoryRecheck', finalInventoryStartedAt);
    if (isWebsiteDirectBooking) {
      logTiming_('publicBookingSubmit:finalInventoryRecheck', finalInventoryStartedAt);
    }
    const remainingMin = Number(liveAvailability.availableRooms || 0);
    if (!validated.allowOverbooking && remainingMin < validated.qtyRooms) {
      throw new Error('Only ' + Math.max(0, remainingMin) + ' room(s) remain for this stay.');
    }
    validated.roomIdentifier = normalizeRoomAssignmentForSave_(validated.roomIdentifier, validated.roomTypeId, {
      bedSetup: validated.bedSetup,
      guests: validated.guests,
      qtyRooms: validated.qtyRooms,
      checkIn: validated.checkIn,
      checkOut: validated.checkOut,
      bookingRows: bookingRows,
      roomIndex: roomIndex
    });
    if (validated.status === BOOKING_STATUS_IN_HOUSE) {
      requireCheckInReadiness_({
        room_identifier: validated.roomIdentifier,
        payment_method: validated.paymentMethod
      });
    }
    if (validated.status === BOOKING_STATUS_CHECKED_OUT && Number(validated.balanceDue || 0) > 0.009) {
      throw new Error('Balance due must be zero before saving a booking as Checked Out.');
    }

    const generateBookingIdStartedAt = Date.now();
    const headerMap = getHeaderMap_(bookingsSheet);
    bookingId = generateBookingId_(bookingsSheet, headerMap, now);
    logCreateManualBookingTiming('createManualBooking:generateBookingId', generateBookingIdStartedAt);

    const buildRowStartedAt = Date.now();
    const bookingRow = {
      booking_id: bookingId,
      created_at: now,
      request_id: validated.requestId,
      source: validated.source,
      source_detail: validated.sourceDetail,
      guest_name: isWebsiteDirectBooking ? sanitizeSheetInput_(validated.guestName) : validated.guestName,
      guest_phone: isWebsiteDirectBooking ? sanitizeSheetInput_(validated.guestPhone) : validated.guestPhone,
      guest_email: isWebsiteDirectBooking ? sanitizeSheetInput_(validated.guestEmail) : validated.guestEmail,
      country: isWebsiteDirectBooking ? sanitizeSheetInput_(validated.country) : validated.country,
      check_in: validated.checkIn,
      check_in_time: validated.checkInTime,
      check_out: validated.checkOut,
      check_out_time: validated.checkOutTime,
      nights: validated.nights,
      room_type_id: validated.roomTypeId,
      room_type_name: validated.roomTypeName,
      room_identifier: validated.roomIdentifier,
      bed_setup: validated.bedSetup,
      adults: validated.adults,
      children: validated.children,
      guests: validated.guests,
      qty_rooms: validated.qtyRooms,
      status: validated.status,
      booking_value: validated.bookingValue,
      booking_value_original: validated.bookingValueOriginal,
      booking_currency: validated.bookingCurrency,
      fx_rate_to_gbp: validated.fxRateToGbp,
      booking_value_gbp: validated.bookingValueGbp,
      pricing_source: validated.pricingSource,
      pricing_reference_id: validated.pricingReferenceId,
      amount_paid: validated.amountPaid,
      payment_method: validated.paymentMethod,
      tax_amount: validated.taxAmount,
      payment_status: validated.paymentStatus,
      payment_notes: validated.paymentNotes,
      balance_due: validated.balanceDue,
      currency: validated.currency,
      cancelled_at: '',
      cancel_reason: '',
      converted_from_request_at: validated.requestId ? now : '',
      guest_preferences: isWebsiteDirectBooking ? sanitizeSheetInput_(validated.guestPreferences) : validated.guestPreferences,
      notes: isWebsiteDirectBooking ? sanitizeSheetInput_(validated.notes) : validated.notes,
      internal_notes: validated.internalNotes
    };
    logCreateManualBookingTiming('createManualBooking:buildRow', buildRowStartedAt);
    const appendStartedAt = Date.now();
    rowNumber = appendObjectRow_(bookingsSheet, headerMap, bookingRow);
    logCreateManualBookingTiming('createManualBooking:appendRow', appendStartedAt);
    logTiming_(isWebsiteDirectBooking ? 'websiteBooking:appendBookingRow' : 'createBooking:appendBookingRow', appendStartedAt);
    if (isWebsiteDirectBooking) {
      logTiming_('publicBookingSubmit:writeBooking', appendStartedAt);
    }
  } finally {
    const lockReleaseStartedAt = Date.now();
    lock.releaseLock();
    logCreateManualBookingTiming('createManualBooking:lockRelease', lockReleaseStartedAt);
  }

  const postCommitConfig = {
    bookingId: bookingId,
    rowNumber: rowNumber,
    sheet: bookingsSheet,
    createdAt: now,
    requestId: validated.requestId,
    source: validated.source,
    bookingValue: validated.bookingValue,
    validated: validated,
    skipRequestSync: validated.skipRequestSync,
    skipRefresh: validated.skipRefresh,
    skipAvailabilityRefresh: validated.skipAvailabilityRefresh,
    internalNotes: validated.internalNotes
  };
  const deferredResultStartedAt = Date.now();
  const postCommit = validated.deferPostCommit
    ? buildDeferredBookingPostCommitResult_(postCommitConfig)
    : runBookingPostCommitTasks_(postCommitConfig);
  logCreateManualBookingTiming('createManualBooking:deferredResult', deferredResultStartedAt);

  const result = {
    ok: true,
    bookingId: bookingId,
    rowNumber: rowNumber,
    bookingCommitted: true,
    source: validated.source,
    bookingValue: validated.bookingValue,
    bookingValueOriginal: validated.bookingValueOriginal,
    bookingCurrency: validated.bookingCurrency,
    fxRateToGbp: validated.fxRateToGbp,
    bookingValueGbp: validated.bookingValueGbp,
    pricingSource: validated.pricingSource,
    pricingReferenceId: validated.pricingReferenceId,
    taxAmount: validated.taxAmount,
    balanceDue: validated.balanceDue,
    paymentStatus: validated.paymentStatus,
    paymentMethod: validated.paymentMethod,
    currency: validated.currency,
    reportingCurrency: validated.reportingCurrency,
    roomTypeId: validated.roomTypeId,
    roomTypeName: validated.roomTypeName,
    roomIdentifier: validated.roomIdentifier,
    checkIn: validated.checkIn,
    checkOut: validated.checkOut,
    nights: validated.nights,
    recoveryNeeded: postCommit.recoveryNeeded,
    recoveryMessage: postCommit.recoveryMessage,
    postCommit: postCommit,
    internalNotes: postCommit.internalNotes
  };
  logCreateManualBookingTiming('createManualBooking:end', createManualBookingStartedAt);
  return result;
}

function cancelBooking(bookingId, reason, options) {
  const cancelStartedAt = Date.now();
  logTiming_('cancel:start');
  const targetBookingId = String(bookingId || '').trim();
  if (!targetBookingId) {
    throw new Error('Booking ID is required.');
  }
  const config = options || {};

  const findBookingStartedAt = Date.now();
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  assertSheetHeadersReady_(sheet, BOOKINGS_HEADERS);
  const rowNumber = findRowNumberByHeaderValue_(sheet, 'booking_id', targetBookingId);

  if (!rowNumber) {
    throw new Error('Booking not found: ' + targetBookingId);
  }

  let existing = getRowObjectByNumber_(sheet, rowNumber);
  logTiming_('cancel:findBooking', findBookingStartedAt);

  const cancelledAt = new Date();
  const cancelReason = String(reason || 'Cancelled manually').trim();
  const lock = LockService.getScriptLock();
  const lockWaitStartedAt = Date.now();
  lock.waitLock(30000);
  logTiming_('cancel:lockWait', lockWaitStartedAt);
  try {
    const lockedExisting = getRowObjectByNumber_(sheet, rowNumber);
    if (!lockedExisting || String(lockedExisting.booking_id || '').trim() !== targetBookingId) {
      throw new Error('Booking changed while cancellation was being prepared. Please refresh and try again.');
    }
    existing = lockedExisting;
    const updateStartedAt = Date.now();
    updateObjectRowBulk_(sheet, rowNumber, {
      status: BOOKING_STATUS_CANCELLED,
      cancelled_at: cancelledAt,
      cancel_reason: cancelReason
    });
    logTiming_('cancel:updateRow', updateStartedAt);
  } finally {
    lock.releaseLock();
  }

  const refreshStartedAt = Date.now();
  logTiming_('cancel:refreshAvailabilityWindow:start');
  const cacheRefresh = refreshAvailabilityCacheWindow_({
    startDate: existing && existing.check_in,
    endDate: existing && existing.check_out,
    roomTypeIds: [String(existing && existing.room_type_id || '').trim()]
  });
  logTiming_('cancel:refreshAvailabilityWindow:end', refreshStartedAt);

  const result = {
    ok: true,
    bookingId: targetBookingId,
    cancelled: true,
    status: BOOKING_STATUS_CANCELLED,
    cancelledAt: cancelledAt,
    cancelReason: cancelReason,
    refreshType: 'availability-window',
    reportingDirty: true,
    availabilityCacheUpdated: cacheRefresh,
    cacheRefreshStatus: cacheRefresh && cacheRefresh.ok ? 'ok' : 'unknown',
    booking: config.includeBookingDetail ? getBookingDetailById_(targetBookingId, { skipGuestHistory: true }) : null
  };
  logTiming_('cancel:return', cancelStartedAt);
  logTiming_('cancel:end', cancelStartedAt);
  return result;
}

function convertRequestToBooking(input) {
  const requestId = String((input && (input.request_id || input.requestId)) || '').trim();
  if (!requestId) {
    throw new Error('Request ID is required.');
  }

  const spreadsheet = getSpreadsheet_();
  const requestSheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.REQUESTS, REQUESTS_HEADERS);
  const requestRowNumber = findRowNumberByHeaderValue_(requestSheet, 'request_id', requestId);

  if (!requestRowNumber) {
    throw new Error('Request not found: ' + requestId);
  }

  const requestRow = getRowObjectByNumber_(requestSheet, requestRowNumber);
  if (String(requestRow.booking_id || '').trim()) {
    throw new Error('Request already linked to booking ' + requestRow.booking_id);
  }

  const bookingResult = createManualBooking(Object.assign({}, input || {}, {
    request_id: requestId,
    guest_name: input && input.guest_name ? input.guest_name : requestRow.guest_name,
    guest_phone: input && input.guest_phone ? input.guest_phone : requestRow.guest_phone,
    guest_email: input && input.guest_email ? input.guest_email : requestRow.guest_email,
    check_in: input && input.check_in ? input.check_in : requestRow.check_in,
    check_out: input && input.check_out ? input.check_out : requestRow.check_out,
    room_type: input && input.room_type ? input.room_type : (requestRow.room_type_id || requestRow.room_type_name),
    bed_setup: input && input.bed_setup ? input.bed_setup : requestRow.bed_setup,
    guests: input && input.guests ? input.guests : requestRow.guests,
    source: input && input.source ? input.source : mapRequestSourceToBookingSource_(requestRow.request_source)
  }));

  return {
    ok: true,
    requestId: requestId,
    booking: bookingResult
  };
}

// -----------------------------------------------------------------------------
// Validation helpers
// -----------------------------------------------------------------------------

function validateAvailabilityInput_(input, options) {
  const config = options || {};
  const checkIn = normalizeDateInput_(input.checkIn || input.check_in);
  const checkOut = normalizeDateInput_(input.checkOut || input.check_out);
  const roomType = String(input.roomType || input.room_type || input.roomTypeId || input.room_type_id || '').trim();
  const guests = Number(input.guests || 0);

  if (!checkIn) throw new Error('Missing check-in date.');
  if (!checkOut) throw new Error('Missing check-out date.');
  if (!roomType) throw new Error('Missing room type.');
  if (!isFinite(guests) || guests < 1 || Math.floor(guests) !== guests) throw new Error('Guests must be a whole number of at least 1.');
  if (checkOut.getTime() <= checkIn.getTime()) throw new Error('Check-out must be after check-in.');
  if (config.publicDateBounds) {
    validatePublicStayDateBounds_(checkIn, checkOut);
  }

  if (!config.skipCapacityValidation) {
    const roomTypeId = resolveRoomTypeId_(roomType);
    const maxGuests = Number(getRoomTypeRowById_(roomTypeId).max_guests || 0);
    if (maxGuests && guests > maxGuests) {
      throw new Error('Selected room type allows up to ' + maxGuests + ' guest(s).');
    }
  }

  return {
    checkIn: checkIn,
    checkOut: checkOut,
    roomType: roomType,
    guests: guests,
    bedSetup: normalizeBedSetup_(input.bedSetup || input.bed_setup || 'Best available')
  };
}

function validatePublicStayDateBounds_(checkIn, checkOut) {
  const today = stripTime_(new Date());
  const checkInDate = stripTime_(checkIn);
  const checkOutDate = stripTime_(checkOut);
  const nights = enumerateStayDates_(checkInDate, checkOutDate).length;
  const latestCheckIn = addDays_(today, PUBLIC_MAX_LOOKAHEAD_DAYS);

  if (checkInDate.getTime() < today.getTime()) {
    throw new Error('Check-in cannot be in the past.');
  }
  if (nights > PUBLIC_MAX_STAY_NIGHTS) {
    throw new Error('Online booking currently supports stays up to ' + PUBLIC_MAX_STAY_NIGHTS + ' nights. Please continue on WhatsApp for longer stays.');
  }
  if (checkInDate.getTime() > latestCheckIn.getTime()) {
    throw new Error('Online booking currently supports check-in dates up to ' + PUBLIC_MAX_LOOKAHEAD_DAYS + ' days ahead. Please continue on WhatsApp for later dates.');
  }

  return {
    nights: nights,
    latestCheckIn: latestCheckIn
  };
}

function validateStaySearchInput_(input) {
  const payload = input || {};
  const checkIn = normalizeDateInput_(payload.checkIn || payload.check_in);
  const checkOut = normalizeDateInput_(payload.checkOut || payload.check_out);
  const guests = Number(payload.guests || 0);
  const requestedBedSetup = normalizeBedSetup_(payload.bedSetup || payload.bed_setup || 'Best available') || 'Best available';
  const effectiveBedSetup = normalizeWebsiteBedSetupPreference_(guests, requestedBedSetup);

  if (!checkIn) throw new Error('Missing check-in date.');
  if (!checkOut) throw new Error('Missing check-out date.');
  if (!guests || guests < 1) throw new Error('Guests must be at least 1.');
  if (guests > 4) throw new Error('Online booking currently supports up to 4 guests. Please continue on WhatsApp for larger groups.');
  if (checkOut.getTime() <= checkIn.getTime()) throw new Error('Check-out must be after check-in.');
  validatePublicStayDateBounds_(checkIn, checkOut);

  return {
    checkIn: checkIn,
    checkOut: checkOut,
    guests: guests,
    requestedBedSetup: requestedBedSetup,
    bedSetup: effectiveBedSetup
  };
}

function getGuestsPerRoom_(guests, qtyRooms) {
  const totalGuests = Math.max(1, Number(guests || 1));
  const roomCount = Math.max(1, Number(qtyRooms || 1));
  return Math.max(1, Math.ceil(totalGuests / roomCount));
}

function validateRequestInput_(input) {
  const payload = input || {};
  ensurePublicSubmissionProtectionApplied_(payload, 'request');
  const availabilityInput = validateAvailabilityInput_(payload, { publicDateBounds: true });
  const guestName = String(payload.guest_name || payload.guestName || 'Website Guest').trim() || 'Website Guest';
  const guestEmail = String(payload.guest_email || payload.guestEmail || '').trim();

  return Object.assign({}, availabilityInput, {
    guest_name: guestName,
    guest_phone: String(payload.guest_phone || payload.guestPhone || '').trim(),
    guest_email: guestEmail,
    request_source: String(payload.request_source || payload.requestSource || REQUEST_SOURCE_WEBSITE).trim().toLowerCase() || REQUEST_SOURCE_WEBSITE,
    notes: String(payload.notes || '').trim()
  });
}

function validateWebsiteBookingInput_(input, options) {
  const payload = input || {};
  const config = options || {};
  const availabilityInput = validateAvailabilityInput_({
    check_in: payload.check_in || payload.checkIn,
    check_out: payload.check_out || payload.checkOut,
    room_type: payload.room_type || payload.roomType || payload.room_type_id || payload.roomTypeId,
    guests: payload.guests,
    bed_setup: payload.bed_setup || payload.bedSetup || 'Best available'
  }, { publicDateBounds: true });
  const guestName = String(payload.guest_name || payload.guestName || '').trim();
  const guestEmail = String(payload.guest_email || payload.guestEmail || '').trim();
  const guestPhone = String(payload.guest_phone || payload.guestPhone || '').trim();
  const country = String(payload.country || payload.guest_country || payload.guestCountry || '').trim();
  const roomTypeId = resolveRoomTypeId_(availabilityInput.roomType);
  const sellableProductId = String(payload.sellable_product_id || payload.sellableProductId || payload.product_id || payload.productId || '').trim();
  if (!sellableProductId) {
    throw new Error('Please select a live room product before booking via website.');
  }
  const sellableProduct = validateSellableProductSelection_(sellableProductId, roomTypeId, availabilityInput.guests);
  const effectiveBedSetup = normalizeWebsiteBedSetupPreference_(availabilityInput.guests, availabilityInput.bedSetup);
  if (!sellableProductSupportsRequestedBedSetup_(sellableProduct, effectiveBedSetup)) {
    throw new Error('Selected room product is not valid for this bed setup.');
  }
  const sellableProductConstraint = getSellableProductConstraint_(sellableProduct, effectiveBedSetup, config.roomIndex);
  const websiteBedSetup = sellableProductConstraint.appliedBedSetup || effectiveBedSetup || 'Best available';

  if (!guestName) throw new Error('Guest full name is required.');
  if (!guestEmail) throw new Error('Guest email is required.');
  if (!isValidEmailAddress_(guestEmail)) throw new Error('Guest email format is invalid.');
  if (!guestPhone) throw new Error('Guest phone number is required.');
  if (!country) throw new Error('Country is required.');

  return {
    guestName: guestName,
    guestEmail: guestEmail,
    guestPhone: guestPhone,
    country: country,
    checkIn: availabilityInput.checkIn,
    checkOut: availabilityInput.checkOut,
    roomTypeId: roomTypeId,
    roomTypeName: getRoomTypeNameById_(roomTypeId),
    sellableProductId: sellableProductId,
    sellableProduct: sellableProduct,
    sellableProductConstraint: sellableProductConstraint,
    sellableProductLabel: sellableProduct ? String(sellableProduct.product_label || sellableProduct.product_name || '').trim() : '',
    guests: Number(availabilityInput.guests || 1),
    bedSetup: websiteBedSetup,
    requestedBedSetup: availabilityInput.bedSetup || 'Best available',
    notes: String(payload.notes || '').trim()
  };
}

function validateManualBookingInput_(input, options) {
  const payload = input || {};
  const config = options || {};
  const existingRow = config.existingRow || null;
  const adults = Math.max(0, Number(payload.adults || 0));
  const children = Math.max(0, Number(payload.children || 0));
  const guestsInput = Number(payload.guests || 0);
  const guests = Math.max(1, guestsInput || adults + children || 1);
  const qtyRooms = Math.max(1, Number(payload.qty_rooms || payload.qtyRooms || 1));
  const guestsPerRoom = getGuestsPerRoom_(guests, qtyRooms);
  const availabilityInput = validateAvailabilityInput_({
    check_in: payload.check_in || payload.checkIn,
    check_out: payload.check_out || payload.checkOut,
    room_type: payload.room_type || payload.roomType || payload.room_type_id || payload.roomTypeId,
    guests: guestsPerRoom,
    bed_setup: payload.bed_setup || payload.bedSetup || 'Best available'
  });
  const roomTypeId = resolveRoomTypeId_(availabilityInput.roomType);
  const roomTypeName = getRoomTypeNameById_(roomTypeId);
  const nights = enumerateStayDates_(availabilityInput.checkIn, availabilityInput.checkOut).length;
  const reportingCurrency = getReportingCurrency_();
  const bookingValueOriginal = validateMoneyAmountInput_(
    payload.booking_value_original != null ? payload.booking_value_original :
      (payload.bookingValueOriginal != null ? payload.bookingValueOriginal :
        (payload.booking_value != null ? payload.booking_value :
          (payload.bookingValue != null ? payload.bookingValue :
            (payload.total_value != null ? payload.total_value : payload.totalValue)))),
    'Booking value',
    0
  );
  const bookingCurrency = normalizeCurrencyCode_(
    payload.booking_currency || payload.bookingCurrency ||
    (existingRow ? getStoredBookingCurrency_(existingRow) : '') ||
    payload.currency ||
    getDefaultBookingCurrency_()
  ) || getDefaultBookingCurrency_();
  const existingOriginalValue = existingRow ? getStoredBookingOriginalValue_(existingRow) : null;
  const existingBookingCurrency = existingRow ? getStoredBookingCurrency_(existingRow) : '';
  const existingFxRate = existingRow ? Number(existingRow.fx_rate_to_gbp || 0) : 0;
  const shouldReuseStoredFx = existingRow &&
    normalizeCurrencyCode_(existingBookingCurrency) === bookingCurrency &&
    Math.abs(roundCurrency_(existingOriginalValue || 0) - bookingValueOriginal) < 0.001 &&
    existingFxRate > 0;
  const suppliedFxRateToGbp = roundFxRate_(payload.fx_rate_to_gbp || payload.fxRateToGbp || payload._fxRateToGbp || 0);
  const skipLiveFx = toBoolean_(payload._skipLiveFx || payload.skipLiveFx);
  const fxRateToGbp = shouldReuseStoredFx
    ? roundFxRate_(existingFxRate)
    : (suppliedFxRateToGbp > 0
      ? suppliedFxRateToGbp
      : (skipLiveFx ? getCachedOrFallbackFxRateToGbp_(bookingCurrency) : getFxRateToGbp_(bookingCurrency)));
  const bookingValueGbp = convertAmountToGbp_(bookingValueOriginal, bookingCurrency, fxRateToGbp);
  assertValidMoneyAmount_(bookingValueGbp, 'Booking value GBP');
  const amountPaid = validateMoneyAmountInput_(
    payload.amount_paid != null ? payload.amount_paid : payload.amountPaid,
    'Amount paid',
    0
  );
  const taxAmount = validateMoneyAmountInput_(
    payload.tax_amount != null ? payload.tax_amount : payload.taxAmount,
    'Tax amount',
    0
  );
  const guestName = String(payload.guest_name || payload.guestName || 'Manual Guest').trim() || 'Manual Guest';
  const source = normalizeBookingSource_(payload.source || payload.booking_source || payload.request_source || 'Manual');
  const sourceDetail = String(payload.source_detail || payload.sourceDetail || '').trim();
  const status = normalizeBookingStatus_(payload.status || BOOKING_STATUS_CONFIRMED);
  if (amountPaid > bookingValueGbp) {
    throw new Error('Amount paid cannot be greater than the booking value.');
  }
  const paymentStatus = normalizePaymentStatus_(payload.payment_status || payload.paymentStatus, bookingValueGbp, amountPaid);
  const paymentMethod = normalizePaymentMethod_(payload.payment_method || payload.paymentMethod);
  if (guestsPerRoom === 3 && roomTypeId !== 'COTTAGE' && ['Double', 'Twin'].indexOf(availabilityInput.bedSetup || '') !== -1) {
    throw new Error('For 3-guest bookings, use Best available or Triple / Family so room assignment can stay valid.');
  }
  if (guestsPerRoom > 3 && normalizeBedSetup_(availabilityInput.bedSetup || '') === 'Triple') {
    throw new Error('Triple / Family setup is not valid for this occupancy. Use Best available instead.');
  }
  const sellableProductId = String(payload.sellable_product_id || payload.sellableProductId || payload.product_id || payload.productId || '').trim();
  const precheckedLiveQuote = payload._precheckedLiveQuote || null;
  const canUsePrecheckedLiveQuote = !!precheckedLiveQuote &&
    toBoolean_(payload._usePrecheckedLiveQuote) &&
    !existingRow &&
    formatDateKey_(normalizeDateInput_(precheckedLiveQuote.checkIn || precheckedLiveQuote.check_in)) === formatDateKey_(availabilityInput.checkIn) &&
    formatDateKey_(normalizeDateInput_(precheckedLiveQuote.checkOut || precheckedLiveQuote.check_out)) === formatDateKey_(availabilityInput.checkOut) &&
    resolveRoomTypeId_(String(precheckedLiveQuote.roomTypeId || precheckedLiveQuote.room_type_id || '')) === roomTypeId &&
    Number(precheckedLiveQuote.guests || 0) === guests &&
    Math.max(1, Number(precheckedLiveQuote.qtyRooms || precheckedLiveQuote.qty_rooms || 1)) === qtyRooms &&
    normalizeBedSetup_(precheckedLiveQuote.bedSetup || precheckedLiveQuote.bed_setup || 'Best available') === normalizeBedSetup_(availabilityInput.bedSetup || 'Best available') &&
    String(precheckedLiveQuote.sellableProductId || precheckedLiveQuote.sellable_product_id || '').trim() === sellableProductId &&
    precheckedLiveQuote.estimatedPrice != null;
  const liveQuoteOptions = {
    qtyRooms: qtyRooms,
    bedSetup: availabilityInput.bedSetup,
    sellableProductId: sellableProductId,
    excludeBookingId: existingRow ? String(existingRow.booking_id || '').trim() : ''
  };
  if (Object.prototype.hasOwnProperty.call(config, 'bookingRows')) liveQuoteOptions.bookingRows = config.bookingRows;
  if (Object.prototype.hasOwnProperty.call(config, 'blockedRows')) liveQuoteOptions.blockedRows = config.blockedRows;
  if (config.roomIndex) liveQuoteOptions.roomIndex = config.roomIndex;

  const liveQuote = canUsePrecheckedLiveQuote
    ? {
        estimatedPrice: roundCurrency_(precheckedLiveQuote.estimatedPrice),
        pricingSource: String(precheckedLiveQuote.pricingSource || precheckedLiveQuote.pricing_source || '').trim(),
        pricingReferenceId: String(precheckedLiveQuote.pricingReferenceId || precheckedLiveQuote.pricing_reference_id || '').trim()
      }
    : buildStayAvailabilityPricingSnapshot_(availabilityInput.checkIn, availabilityInput.checkOut, roomTypeId, guests, liveQuoteOptions);
  const existingPricingSource = existingRow ? String(existingRow.pricing_source || '').trim() : '';
  const existingPricingReferenceId = existingRow ? String(existingRow.pricing_reference_id || '').trim() : '';
  const pricingInputsChanged = existingRow &&
    (formatDateKey_(normalizeDateInput_(existingRow.check_in)) !== formatDateKey_(availabilityInput.checkIn) ||
      formatDateKey_(normalizeDateInput_(existingRow.check_out)) !== formatDateKey_(availabilityInput.checkOut) ||
      resolveRoomTypeId_(String(existingRow.room_type_id || existingRow.roomTypeId || '').trim()) !== roomTypeId ||
      Math.max(1, Number(existingRow.guests || Number(existingRow.adults || 0) + Number(existingRow.children || 0) || 1)) !== guests ||
      Math.max(1, Number(existingRow.qty_rooms || 1)) !== qtyRooms);
  const retainedPricingSource = pricingInputsChanged ? '' : existingPricingSource;
  const retainedPricingReferenceId = pricingInputsChanged ? '' : existingPricingReferenceId;
  const quoteMatchesInput = liveQuote.estimatedPrice != null && Math.abs(roundCurrency_(liveQuote.estimatedPrice) - bookingValueOriginal) < 0.01;
  const pricingSource = String(payload.pricing_source || payload.pricingSource || '').trim() ||
    (quoteMatchesInput ? String(liveQuote.pricingSource || '').trim() : '') ||
    retainedPricingSource ||
    'manual_override';
  const pricingReferenceId = String(payload.pricing_reference_id || payload.pricingReferenceId || '').trim() ||
    (quoteMatchesInput ? String(liveQuote.pricingReferenceId || '').trim() : '') ||
    retainedPricingReferenceId ||
    'manual';

  return {
    requestId: String(payload.request_id || payload.requestId || '').trim(),
    sellableProductId: sellableProductId,
    guestName: guestName,
    guestPhone: String(payload.guest_phone || payload.guestPhone || '').trim(),
    guestEmail: String(payload.guest_email || payload.guestEmail || '').trim(),
    country: String(payload.country || payload.guest_country || payload.guestCountry || '').trim(),
    checkIn: availabilityInput.checkIn,
    checkInTime: normalizeTimeInput_(payload.check_in_time || payload.checkInTime),
    checkOut: availabilityInput.checkOut,
    checkOutTime: normalizeTimeInput_(payload.check_out_time || payload.checkOutTime),
    nights: nights,
    roomTypeId: roomTypeId,
    roomTypeName: roomTypeName,
    roomIdentifier: String(payload.room_identifier || payload.roomIdentifier || '').trim(),
    bedSetup: availabilityInput.bedSetup || 'Best available',
    adults: adults || guests,
    children: children,
    guests: guests,
    qtyRooms: qtyRooms,
    source: source,
    sourceDetail: sourceDetail,
    status: status,
    bookingValue: bookingValueGbp,
    bookingValueOriginal: bookingValueOriginal,
    bookingCurrency: bookingCurrency,
    fxRateToGbp: fxRateToGbp,
    bookingValueGbp: bookingValueGbp,
    pricingSource: pricingSource,
    pricingReferenceId: pricingReferenceId,
    amountPaid: amountPaid,
    paymentMethod: paymentMethod,
    taxAmount: taxAmount,
    paymentStatus: paymentStatus,
    paymentNotes: String(payload.payment_notes || payload.paymentNotes || '').trim(),
    balanceDue: calculateBalanceDue_(bookingValueGbp, amountPaid),
    currency: reportingCurrency,
    reportingCurrency: reportingCurrency,
    guestPreferences: String(payload.guest_preferences || payload.guestPreferences || '').trim(),
    notes: String(payload.notes || '').trim(),
    internalNotes: String(payload.internal_notes || payload.internalNotes || '').trim(),
    allowOverbooking: toBoolean_(payload.allow_overbooking || payload.allowOverbooking),
    skipRequestSync: toBoolean_(payload._skipRequestSync),
    skipRefresh: toBoolean_(payload._skipRefresh),
    skipAvailabilityRefresh: toBoolean_(payload._skipAvailabilityRefresh || payload.skipAvailabilityRefresh),
    skipStructureEnsure: toBoolean_(payload._skipStructureEnsure || payload.skipStructureEnsure),
    fastInventoryRows: toBoolean_(payload._fastInventoryRows || payload.fastInventoryRows),
    deferPostCommit: toBoolean_(payload._deferPostCommit || payload.deferPostCommit)
  };
}

function validateBlockInput_(input) {
  const payload = input || {};
  const startDate = normalizeDateInput_(payload.start_date || payload.startDate);
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate);
  const roomTypeRaw = String(payload.room_type || payload.roomType || payload.room_type_id || payload.roomTypeId || '').trim();
  const qtyBlocked = Math.max(1, Number(payload.qty_blocked || payload.qtyBlocked || payload.quantity || 1));

  if (!startDate) throw new Error('Block start date is required.');
  if (!endDate) throw new Error('Block end date is required.');
  if (!roomTypeRaw) throw new Error('Room type is required for a block.');
  if (endDate.getTime() < startDate.getTime()) throw new Error('Block end date must be on or after the start date.');

  const roomTypeId = resolveRoomTypeId_(roomTypeRaw);
  const inventoryTotal = getRoomInventory_(roomTypeId);
  if (inventoryTotal > 0 && qtyBlocked > inventoryTotal) {
    throw new Error('Only ' + inventoryTotal + ' room(s) exist for ' + getRoomTypeNameById_(roomTypeId) + '.');
  }

  return {
    startDate: startDate,
    endDate: endDate,
    roomTypeId: roomTypeId,
    roomTypeName: getRoomTypeNameById_(roomTypeId),
    qtyBlocked: qtyBlocked,
    reason: String(payload.reason || '').trim() || 'Manual block',
    notes: String(payload.notes || '').trim()
  };
}

function buildAvailabilityBreakdown_(checkIn, checkOut, roomTypeId, options) {
  const config = options || {};
  const stayDates = enumerateStayDates_(checkIn, checkOut);
  const physicalInventory = Number(config.inventoryTotal != null ? config.inventoryTotal : getRoomInventory_(roomTypeId));
  const requestedGuests = Math.max(0, Number(config.guests || 0));
  const requestedBedSetup = normalizeBedSetup_(config.bedSetup || config.bed_setup || '');
  const requestedCandidateRoomIds = Array.isArray(config.candidateRoomIds || config.requestedCandidateRoomIds)
    ? (config.candidateRoomIds || config.requestedCandidateRoomIds).map(function(value) {
        return String(value || '').trim();
      }).filter(Boolean)
    : [];
  const needsCompatibleInventory = requestedGuests > 0 || requestedBedSetup || requestedCandidateRoomIds.length || String(config.excludeBookingId || '').trim();

  if (!needsCompatibleInventory) {
    const cacheRows = (config.cacheRows || getSheetObjects_(SHEET_NAMES.AVAILABILITY_CACHE)).filter(function(row) {
      return String(row.room_type_id || '').trim() === roomTypeId;
    });
    const cacheMap = {};

    cacheRows.forEach(function(row) {
      const date = normalizeDateInput_(row.date);
      if (!date) return;
      cacheMap[formatDateKey_(date)] = row;
    });

    const cacheComplete = stayDates.every(function(date) {
      return Object.prototype.hasOwnProperty.call(cacheMap, formatDateKey_(date));
    });

    if (cacheComplete) {
      return stayDates.map(function(date) {
        const key = formatDateKey_(date);
        const row = cacheMap[key];
        return {
          date: key,
          totalRooms: Number(row.inventory_total || physicalInventory),
          soldRooms: Number(row.booked_confirmed || 0),
          blockedRooms: Number(row.blocked || 0),
          availableRooms: Number(row.remaining || 0),
          status: row.status || (Number(row.remaining || 0) < 0 ? 'Overbooked' : (Number(row.remaining || 0) === 0 ? 'Sold Out' : 'Available')),
          source: 'cache'
        };
      });
    }
  }

  const mapBounds = {
    startDate: checkIn,
    endDateExclusive: checkOut
  };
  const blockedCountMap = config.blockedCountMap || buildBlockedDateQtyMap_(config.blockedRows || getSheetObjects_(SHEET_NAMES.BLOCKED_DATES), mapBounds).byDateRoom;

  if (needsCompatibleInventory) {
    const roomIndex = config.roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
    return stayDates.map(function(date) {
      const compatible = getCompatibleAvailabilityForDate_(date, roomTypeId, Math.max(1, requestedGuests || 1), requestedBedSetup || 'Best available', {
        bookingRows: config.bookingRows,
        blockedCountMap: blockedCountMap,
        demandUnitsByDateRoom: config.demandUnitsByDateRoom,
        roomIndex: roomIndex,
        requestedCandidateRoomIds: requestedCandidateRoomIds,
        excludeBookingId: config.excludeBookingId
      });
      return {
        date: formatDateKey_(date),
        totalRooms: compatible.physicalInventory,
        soldRooms: compatible.soldRooms,
        blockedRooms: compatible.blockedRooms,
        availableRooms: compatible.availableRooms,
        compatibleInventory: compatible.compatibleInventory,
        status: compatible.availableRooms < 0 ? 'Overbooked' : (compatible.availableRooms === 0 ? 'Sold Out' : 'Available'),
        source: 'live'
      };
    });
  }

  const bookingCountMap = config.bookingCountMap || buildConfirmedBookingRoomCountMap_(config.bookingRows || getSheetObjects_(SHEET_NAMES.BOOKINGS), mapBounds).byDateRoom;
  const commercialControls = config.commercialControls || getCommercialControlRows_({ activeOnly: true, roomTypeId: roomTypeId });
  return stayDates.map(function(date) {
    const totalRooms = getSellableInventoryTotalForDate_(date, roomTypeId, commercialControls, physicalInventory);
    const soldRooms = getConfirmedBookingsCount_(date, roomTypeId, bookingCountMap);
    const blockedRooms = getBlockedDatesCount_(date, roomTypeId, blockedCountMap);
    const availableRooms = totalRooms - soldRooms - blockedRooms;
    return {
      date: formatDateKey_(date),
      totalRooms: totalRooms,
      soldRooms: soldRooms,
      blockedRooms: blockedRooms,
      availableRooms: availableRooms,
      status: availableRooms < 0 ? 'Overbooked' : (availableRooms === 0 ? 'Sold Out' : 'Available'),
      source: 'live'
    };
  });
}

function buildFrontDeskDashboardData_(selectedDate, options) {
  const config = options || {};
  const date = normalizeFrontDeskDate_(selectedDate);
  const today = stripTime_(new Date());
  const fastOperational = toBoolean_(config.fastOperational || config.fastMode);
  const skipAvailabilitySnapshot = toBoolean_(config.skipAvailabilitySnapshot || config.skip_availability_snapshot);
  const reportingCurrency = config.reportingCurrency || getReportingCurrency_();
  const defaultBookingCurrency = config.defaultBookingCurrency || getDefaultBookingCurrency_();
  const bookings = config.bookings || getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const bookingsAreOperationalSubset = toBoolean_(config.bookingsAreOperationalSubset || config.operationalBookingSubset);
  const roomIndex = config.roomIndex || buildRoomMasterIndex_(config.roomRows || getRoomsMasterRows_({ activeOnly: false }));
  const bookingListOptions = {
    bookings: bookings,
    roomIndex: roomIndex,
    reportingCurrency: reportingCurrency,
    defaultBookingCurrency: defaultBookingCurrency,
    skipGuestHistory: config.staffMode || config.skipGuestHistory
  };
  if (Object.prototype.hasOwnProperty.call(config, 'historyRows') && config.historyRows !== null && config.historyRows !== undefined) {
    bookingListOptions.historyRows = config.historyRows;
  }
  const bookingLists = buildFrontDeskBookingLists_(date, bookingListOptions);
  const arrivals = bookingLists.arrivals;
  const departures = bookingLists.departures;
  const inHouse = bookingLists.inHouse;
  if (config.staffMode) {
    return {
      ok: true,
      selectedDate: formatDateKey_(date),
      arrivals: arrivals,
      departures: departures,
      inHouse: inHouse,
      counts: {
        arrivals: arrivals.length,
        departures: departures.length,
        inHouse: inHouse.length
      }
    };
  }
  const roomTypes = skipAvailabilitySnapshot ? [] : (config.roomTypes || getActiveRoomTypeCatalog_());
  const availabilityCacheRows = skipAvailabilitySnapshot ? [] : getSheetObjects_(SHEET_NAMES.AVAILABILITY_CACHE);
  const availability = skipAvailabilitySnapshot
    ? {
      ok: true,
      selectedDate: formatDateKey_(date),
      source: 'deferred',
      roomTypes: [],
      totals: {
        totalRooms: 0,
        soldRooms: 0,
        blockedRooms: 0,
        availableRooms: 0
      }
    }
    : buildAvailabilitySnapshotForDate_(date, '', {
      bookings: bookings,
      availabilityCacheRows: availabilityCacheRows,
      roomTypes: roomTypes
    });
  const monthStart = startOfMonth_(date);
  const insights = fastOperational
    ? {
      directWebsiteSharePct: 0,
      directWebsiteBookings: 0,
      cancellationsCount: 0,
      averageLeadTimeDays: 0
    }
    : buildPortfolioInsights_(bookings, monthStart, date, {
      defaultBookingCurrency: defaultBookingCurrency
    });
  const tonightAvailability = skipAvailabilitySnapshot
    ? availability
    : formatDateKey_(date) === formatDateKey_(today)
    ? availability
    : buildAvailabilitySnapshotForDate_(today, '', {
    bookings: bookings,
    availabilityCacheRows: availabilityCacheRows,
    roomTypes: roomTypes
  });
  const trend = fastOperational || skipAvailabilitySnapshot ? [] : buildAvailabilityCacheOccupancyTrend_(today, 7, availabilityCacheRows, roomTypes);

  return {
    ok: true,
    selectedDate: formatDateKey_(date),
    arrivals: arrivals,
    departures: departures,
    inHouse: inHouse,
    availability: availability,
    metrics: {
      occupancyTonightPct: safeDivide_(Number(tonightAvailability.totals.soldRooms || 0), Number(tonightAvailability.totals.totalRooms || 0)),
      occupancyTonightSold: Number(tonightAvailability.totals.soldRooms || 0),
      occupancyTonightTotal: Number(tonightAvailability.totals.totalRooms || 0),
      unassignedArrivals: countUnassignedBookings_(arrivals),
      repeatArrivals: countRepeatGuests_(arrivals),
      unpaidDepartures: countBalanceDueBookings_(departures),
      unpaidDeparturesBalance: sumBookingBalanceDue_(departures),
      directWebsiteSharePct: insights.directWebsiteSharePct,
      directWebsiteBookings: insights.directWebsiteBookings,
      cancellationsMtd: insights.cancellationsCount,
      averageLeadTimeDays: insights.averageLeadTimeDays
    },
    nextSevenDayTrend: trend,
    deferred: {
      availabilitySnapshot: skipAvailabilitySnapshot,
      guestHistory: !!config.skipGuestHistory,
      portfolioInsights: fastOperational
    },
    counts: {
      arrivals: arrivals.length,
      departures: departures.length,
      inHouse: inHouse.length,
      availableRooms: availability.totals.availableRooms
    }
  };
}

function buildFrontDeskBookingLists_(selectedDate, options) {
  const config = options || {};
  const date = normalizeFrontDeskDate_(selectedDate);
  const rows = config.bookings || getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const defaultBookingCurrency = config.defaultBookingCurrency || getDefaultBookingCurrency_();
  const roomIndex = config.roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const arrivalRawRows = rows.filter(function(row) {
    return matchesFrontDeskBookingMode_(row, date, 'arrivals');
  });
  const departureRawRows = rows.filter(function(row) {
    return matchesFrontDeskBookingMode_(row, date, 'departures');
  });
  const inHouseRawRows = rows.filter(function(row) {
    return matchesFrontDeskBookingMode_(row, date, 'inHouse');
  });
  const targetBookingIds = arrivalRawRows.concat(departureRawRows, inHouseRawRows).map(function(row) {
    return String(row.booking_id || '').trim();
  }).filter(Boolean);
  const historyRows = Object.prototype.hasOwnProperty.call(config, 'historyRows')
    ? config.historyRows
    : rows;
  const historyIndex = config.skipGuestHistory
    ? null
    : (config.historyIndex || buildGuestHistoryIndex_(historyRows, {
      defaultBookingCurrency: defaultBookingCurrency,
      targetBookingIds: targetBookingIds
    }));
  const displayContext = {
    historyIndex: historyIndex,
    roomIndex: roomIndex,
    reportingCurrency: config.reportingCurrency || getReportingCurrency_(),
    defaultBookingCurrency: defaultBookingCurrency
  };

  function mapDisplay(row) {
    return mapBookingRowToAdminDisplay_(row, displayContext);
  }

  const arrivalRows = arrivalRawRows.map(mapDisplay).sort(function(a, b) {
    return compareArrivalDisplayRows_(a, b);
  });

  const departureRows = departureRawRows.sort(function(a, b) {
    return compareFrontDeskBookings_(a, b, 'departures');
  }).map(mapDisplay);

  const inHouseRows = inHouseRawRows.sort(function(a, b) {
    return compareFrontDeskBookings_(a, b, 'inHouse');
  }).map(mapDisplay);

  return {
    arrivals: arrivalRows,
    departures: departureRows,
    inHouse: inHouseRows
  };
}

function buildAvailabilityCacheOccupancyTrend_(startDate, days, availabilityCacheRows, roomTypes) {
  const safeDays = Math.max(1, Number(days || 7));
  const types = (roomTypes || getActiveRoomTypeCatalog_()).filter(function(roomType) {
    return roomType.roomTypeId;
  });
  if (!types.length) return [];

  const requestedIds = types.map(function(roomType) {
    return roomType.roomTypeId;
  });
  const requestedLookup = requestedIds.reduce(function(map, roomTypeId) {
    map[roomTypeId] = true;
    return map;
  }, {});
  const rangeLookup = {};
  const rowsByDateRoom = {};

  for (let offset = 0; offset < safeDays; offset++) {
    rangeLookup[formatDateKey_(addDays_(startDate, offset))] = true;
  }

  (availabilityCacheRows || getSheetObjects_(SHEET_NAMES.AVAILABILITY_CACHE)).forEach(function(row) {
    const cacheDate = normalizeDateInput_(row.date);
    const roomTypeId = String(row.room_type_id || '').trim();
    if (!cacheDate || !requestedLookup[roomTypeId]) return;
    const dateKey = formatDateKey_(cacheDate);
    if (!rangeLookup[dateKey]) return;
    rowsByDateRoom[dateKey + '|' + roomTypeId] = row;
  });

  return Object.keys(rangeLookup).sort().map(function(dateKey) {
    const complete = requestedIds.every(function(roomTypeId) {
      return !!rowsByDateRoom[dateKey + '|' + roomTypeId];
    });
    if (!complete) return null;

    const totals = requestedIds.reduce(function(acc, roomTypeId) {
      const row = rowsByDateRoom[dateKey + '|' + roomTypeId] || {};
      acc.totalRooms += Number(row.inventory_total || 0);
      acc.soldRooms += Number(row.booked_confirmed || 0);
      acc.availableRooms += Math.max(0, Number(row.inventory_total || 0) - Number(row.blocked || 0));
      return acc;
    }, {
      totalRooms: 0,
      soldRooms: 0,
      availableRooms: 0
    });

    return {
      date: dateKey,
      occupancyPct: safeDivide_(totals.soldRooms, totals.availableRooms),
      soldRooms: totals.soldRooms,
      availableRooms: totals.availableRooms
    };
  }).filter(Boolean);
}

function getFrontDeskBookingsForDate_(selectedDate, mode) {
  const date = normalizeFrontDeskDate_(selectedDate);
  const rows = getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const defaultBookingCurrency = getDefaultBookingCurrency_();
  const reportingCurrency = getReportingCurrency_();
  const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const filteredRows = rows.filter(function(row) {
    return matchesFrontDeskBookingMode_(row, date, mode);
  });
  const historyIndex = buildGuestHistoryIndex_(rows, {
    defaultBookingCurrency: defaultBookingCurrency,
    targetBookingIds: filteredRows.map(function(row) {
      return String(row.booking_id || '').trim();
    }).filter(Boolean)
  });

  return filteredRows
    .sort(function(a, b) {
      return compareFrontDeskBookings_(a, b, mode);
    })
    .map(function(row) {
      return mapBookingRowToAdminDisplay_(row, {
        historyIndex: historyIndex,
        roomIndex: roomIndex,
        reportingCurrency: reportingCurrency,
        defaultBookingCurrency: defaultBookingCurrency
      });
    });
}

function matchesFrontDeskBookingMode_(row, selectedDate, mode) {
  const status = normalizeBookingStatus_(row.status);
  const checkIn = normalizeDateInput_(row.check_in);
  const checkOut = normalizeDateInput_(row.check_out);
  if (!checkIn || !checkOut) return false;

  const selectedTime = stripTime_(selectedDate).getTime();
  const checkInTime = stripTime_(checkIn).getTime();
  const checkOutTime = stripTime_(checkOut).getTime();
  const normalizedMode = String(mode || '').toLowerCase();

  if (normalizedMode === 'arrivals') {
    return status === BOOKING_STATUS_CONFIRMED && checkInTime === selectedTime;
  }
  if (normalizedMode === 'departures') {
    return status === BOOKING_STATUS_IN_HOUSE && checkOutTime === selectedTime;
  }
  if (normalizedMode === 'inhouse') {
    return status === BOOKING_STATUS_IN_HOUSE && checkInTime <= selectedTime && checkOutTime > selectedTime;
  }
  return false;
}

function compareFrontDeskBookings_(a, b, mode) {
  const normalizedMode = String(mode || '').toLowerCase();
  const useDepartureSort = normalizedMode === 'departures';
  const aPrimary = useDepartureSort ? normalizeDateInput_(a.check_out) : normalizeDateInput_(a.check_in);
  const bPrimary = useDepartureSort ? normalizeDateInput_(b.check_out) : normalizeDateInput_(b.check_in);
  const aPrimaryTime = aPrimary ? aPrimary.getTime() : 0;
  const bPrimaryTime = bPrimary ? bPrimary.getTime() : 0;

  if (aPrimaryTime !== bPrimaryTime) return aPrimaryTime - bPrimaryTime;
  const aSecondary = useDepartureSort ? normalizeTimeInput_(a.check_out_time) : normalizeTimeInput_(a.check_in_time);
  const bSecondary = useDepartureSort ? normalizeTimeInput_(b.check_out_time) : normalizeTimeInput_(b.check_in_time);
  if (aSecondary !== bSecondary) {
    if (!aSecondary) return 1;
    if (!bSecondary) return -1;
    return aSecondary.localeCompare(bSecondary);
  }
  return String(a.guest_name || '').localeCompare(String(b.guest_name || ''));
}

function getArrivalPrepStatusKey_(input) {
  const payload = input || {};
  if (!String(payload.roomIdentifier || '').trim()) return 'needs_room';
  if (payload.hasBedSetupMismatch || payload.hasAttentionNotes || payload.hasPreferences) return 'needs_review';
  if (payload.hasBalanceDue) return 'needs_payment';
  return 'ready';
}

function getArrivalPrepStatusLabel_(statusKey) {
  return {
    needs_room: 'Needs Room',
    needs_review: 'Needs Review',
    needs_payment: 'Needs Payment Follow-up',
    ready: 'Ready'
  }[String(statusKey || '').trim().toLowerCase()] || 'Needs Review';
}

function getArrivalPrepPriority_(statusKey) {
  return {
    needs_room: 0,
    needs_review: 1,
    needs_payment: 2,
    ready: 3
  }[String(statusKey || '').trim().toLowerCase()];
}

function compareArrivalDisplayRows_(a, b) {
  const aCheckIn = normalizeDateInput_(a && a.checkIn ? a.checkIn : '');
  const bCheckIn = normalizeDateInput_(b && b.checkIn ? b.checkIn : '');
  const dateDelta = (aCheckIn ? aCheckIn.getTime() : 0) - (bCheckIn ? bCheckIn.getTime() : 0);
  if (dateDelta !== 0) return dateDelta;

  const aPriority = Number(a && a.prepPriority != null ? a.prepPriority : 99);
  const bPriority = Number(b && b.prepPriority != null ? b.prepPriority : 99);
  if (aPriority !== bPriority) return aPriority - bPriority;

  const aTime = String(a && a.checkInTime ? a.checkInTime : '').trim();
  const bTime = String(b && b.checkInTime ? b.checkInTime : '').trim();
  if (aTime !== bTime) {
    if (!aTime) return 1;
    if (!bTime) return -1;
    return aTime.localeCompare(bTime);
  }

  const guestDelta = String(a && a.guestName ? a.guestName : '').localeCompare(String(b && b.guestName ? b.guestName : ''));
  if (guestDelta !== 0) return guestDelta;
  return String(a && a.bookingId ? a.bookingId : '').localeCompare(String(b && b.bookingId ? b.bookingId : ''));
}

function buildAvailabilitySnapshotForDate_(selectedDate, roomTypeId, options) {
  const config = options || {};
  const date = normalizeFrontDeskDate_(selectedDate);
  const dateKey = formatDateKey_(date);
  const targetRoomTypeId = roomTypeId ? resolveRoomTypeId_(roomTypeId) : '';
  const roomTypes = (config.roomTypes || getActiveRoomTypeCatalog_())
    .filter(function(row) {
      return row.roomTypeId && (!targetRoomTypeId || row.roomTypeId === targetRoomTypeId);
    });

  const requestedIds = roomTypes.map(function(roomType) {
    return roomType.roomTypeId;
  });
  const cacheRows = (config.availabilityCacheRows || getSheetObjects_(SHEET_NAMES.AVAILABILITY_CACHE)).filter(function(row) {
    const cacheDate = normalizeDateInput_(row.date);
    const cacheRoomTypeId = String(row.room_type_id || '').trim();
    return cacheDate &&
      formatDateKey_(cacheDate) === dateKey &&
      requestedIds.indexOf(cacheRoomTypeId) !== -1;
  });
  const cacheMap = {};

  cacheRows.forEach(function(row) {
    cacheMap[String(row.room_type_id || '').trim()] = row;
  });

  const cacheComplete = requestedIds.length > 0 && requestedIds.every(function(id) {
    return Object.prototype.hasOwnProperty.call(cacheMap, id);
  });

  const bookingCountMap = cacheComplete
    ? null
    : (config.bookingCountMap || buildConfirmedBookingRoomCountMap_(config.bookings || getSheetObjects_(SHEET_NAMES.BOOKINGS), {
      startDate: date,
      endDateExclusive: addDays_(date, 1)
    }).byDateRoom);
  const blockedCountMap = cacheComplete
    ? null
    : (config.blockedCountMap || buildBlockedDateQtyMap_(config.blockedRows || getSheetObjects_(SHEET_NAMES.BLOCKED_DATES), {
      startDate: date,
      endDateExclusive: addDays_(date, 1)
    }).byDateRoom);
  const commercialControls = cacheComplete
    ? null
    : (config.commercialControls || getCommercialControlRows_({ activeOnly: true }));

  const rows = roomTypes.map(function(roomType) {
    const row = cacheMap[roomType.roomTypeId];
    const soldRooms = cacheComplete
      ? Number(row.booked_confirmed || 0)
      : getConfirmedBookingsCount_(date, roomType.roomTypeId, bookingCountMap);
    const blockedRooms = cacheComplete
      ? Number(row.blocked || 0)
      : getBlockedDatesCount_(date, roomType.roomTypeId, blockedCountMap);
    const totalRooms = cacheComplete
      ? Number(row.inventory_total || roomType.inventoryTotal)
      : getSellableInventoryTotalForDate_(date, roomType.roomTypeId, commercialControls, roomType.inventoryTotal);
    const availableRooms = cacheComplete
      ? Number(row.remaining || 0)
      : (totalRooms - soldRooms - blockedRooms);

    return {
      roomTypeId: roomType.roomTypeId,
      roomTypeName: roomType.roomTypeName,
      totalRooms: totalRooms,
      soldRooms: soldRooms,
      blockedRooms: blockedRooms,
      availableRooms: availableRooms,
      status: availableRooms < 0 ? 'Overbooked' : (availableRooms === 0 ? 'Sold Out' : 'Available'),
      source: cacheComplete ? 'cache' : 'live'
    };
  });

  const totals = rows.reduce(function(acc, row) {
    acc.totalRooms += Number(row.totalRooms || 0);
    acc.soldRooms += Number(row.soldRooms || 0);
    acc.blockedRooms += Number(row.blockedRooms || 0);
    acc.availableRooms += Number(row.availableRooms || 0);
    return acc;
  }, {
    totalRooms: 0,
    soldRooms: 0,
    blockedRooms: 0,
    availableRooms: 0
  });

  return {
    ok: true,
    selectedDate: dateKey,
    source: rows.length && rows[0].source ? rows[0].source : 'live',
    roomTypes: rows,
    totals: totals
  };
}

function buildAvailabilityRangeOverview_(checkIn, checkOut, options) {
  const config = options || {};
  const roomTypes = (config.roomTypes || getActiveRoomTypeCatalog_())
    .filter(function(row) {
      return row.roomTypeId;
    });
  const cacheRows = config.availabilityCacheRows || config.cacheRows || getSheetObjects_(SHEET_NAMES.AVAILABILITY_CACHE);
  const stayDates = enumerateStayDates_(checkIn, checkOut);
  const stayDateLookup = {};
  stayDates.forEach(function(stayDate) {
    stayDateLookup[formatDateKey_(stayDate)] = true;
  });
  const cacheRelevant = cacheRows.filter(function(row) {
    const cacheDate = normalizeDateInput_(row.date);
    if (!cacheDate) return false;
    return !!stayDateLookup[formatDateKey_(cacheDate)];
  });
  const cacheRelevantMap = {};
  cacheRelevant.forEach(function(row) {
    const cacheDate = normalizeDateInput_(row.date);
    const cacheRoomTypeId = String(row.room_type_id || '').trim();
    if (!cacheDate || !cacheRoomTypeId) return;
    cacheRelevantMap[formatDateKey_(cacheDate) + '|' + cacheRoomTypeId] = true;
  });
  const requestedIds = roomTypes.map(function(roomType) {
    return roomType.roomTypeId;
  });
  const cacheComplete = requestedIds.every(function(roomTypeId) {
    return stayDates.every(function(stayDate) {
      return !!cacheRelevantMap[formatDateKey_(stayDate) + '|' + roomTypeId];
    });
  });
  const bookingRows = cacheComplete ? null : (config.bookingRows || getSheetObjects_(SHEET_NAMES.BOOKINGS));
  const blockedRows = cacheComplete ? null : (config.blockedRows || getSheetObjects_(SHEET_NAMES.BLOCKED_DATES));
  const commercialControls = cacheComplete ? null : (config.commercialControls || getCommercialControlRows_({ activeOnly: true }));
  const mapBounds = {
    startDate: checkIn,
    endDateExclusive: checkOut
  };
  const bookingCountMap = cacheComplete ? null : buildConfirmedBookingRoomCountMap_(bookingRows, mapBounds).byDateRoom;
  const blockedCountMap = cacheComplete ? null : buildBlockedDateQtyMap_(blockedRows, mapBounds).byDateRoom;

  const roomTypeBreakdown = roomTypes.map(function(roomType) {
    const breakdown = buildAvailabilityBreakdown_(checkIn, checkOut, roomType.roomTypeId, {
      inventoryTotal: roomType.inventoryTotal,
      cacheRows: cacheRelevant,
      bookingRows: bookingRows,
      blockedRows: blockedRows,
      bookingCountMap: bookingCountMap,
      blockedCountMap: blockedCountMap,
      commercialControls: commercialControls
    });
    const totalRooms = breakdown.reduce(function(max, row) {
      return Math.max(max, Number(row.totalRooms || roomType.inventoryTotal || 0));
    }, Number(roomType.inventoryTotal || 0));
    const soldRooms = breakdown.reduce(function(max, row) {
      return Math.max(max, Number(row.soldRooms || 0));
    }, 0);
    const blockedRooms = breakdown.reduce(function(max, row) {
      return Math.max(max, Number(row.blockedRooms || 0));
    }, 0);
    const availableRooms = breakdown.reduce(function(min, row) {
      return Math.min(min, Number(row.availableRooms || 0));
    }, roomType.inventoryTotal);

    return {
      roomTypeId: roomType.roomTypeId,
      roomTypeName: roomType.roomTypeName,
      totalRooms: totalRooms,
      soldRooms: soldRooms,
      blockedRooms: blockedRooms,
      availableRooms: availableRooms,
      status: availableRooms < 0 ? 'Overbooked' : (availableRooms === 0 ? 'Sold Out' : 'Available'),
      dailyBreakdown: breakdown
    };
  });

  const dailyBreakdown = stayDates.map(function(stayDate, index) {
    const totals = roomTypeBreakdown.reduce(function(acc, roomType) {
      const day = roomType.dailyBreakdown[index] || {};
      acc.totalRooms += Number(day.totalRooms || roomType.totalRooms || 0);
      acc.soldRooms += Number(day.soldRooms || 0);
      acc.blockedRooms += Number(day.blockedRooms || 0);
      acc.availableRooms += Number(day.availableRooms || 0);
      return acc;
    }, {
      totalRooms: 0,
      soldRooms: 0,
      blockedRooms: 0,
      availableRooms: 0
    });

    return {
      date: formatDateKey_(stayDate),
      totalRooms: totals.totalRooms,
      soldRooms: totals.soldRooms,
      blockedRooms: totals.blockedRooms,
      availableRooms: totals.availableRooms,
      status: totals.availableRooms < 0 ? 'Overbooked' : (totals.availableRooms === 0 ? 'Sold Out' : 'Available')
    };
  });

  const source = roomTypeBreakdown.some(function(row) {
    return (row.dailyBreakdown || []).some(function(day) {
      return day.source === 'live';
    });
  }) ? 'live' : 'cache';

  const result = {
    ok: true,
    checkIn: formatDateKey_(checkIn),
    checkOut: formatDateKey_(checkOut),
    nights: stayDates.length,
    totalRooms: dailyBreakdown.reduce(function(max, row) {
      return Math.max(max, Number(row.totalRooms || 0));
    }, 0),
    soldRooms: dailyBreakdown.reduce(function(max, row) {
      return Math.max(max, Number(row.soldRooms || 0));
    }, 0),
    blockedRooms: dailyBreakdown.reduce(function(max, row) {
      return Math.max(max, Number(row.blockedRooms || 0));
    }, 0),
    availableRooms: dailyBreakdown.reduce(function(min, row) {
      return Math.min(min, Number(row.availableRooms || 0));
    }, dailyBreakdown.length ? Number(dailyBreakdown[0].availableRooms || 0) : 0),
    available: dailyBreakdown.every(function(row) {
      return Number(row.availableRooms || 0) > 0;
    }),
    source: source,
    dailyBreakdown: dailyBreakdown,
    roomTypeBreakdown: roomTypeBreakdown.map(function(row) {
      return {
        roomTypeId: row.roomTypeId,
        roomTypeName: row.roomTypeName,
        totalRooms: row.totalRooms,
        soldRooms: row.soldRooms,
        blockedRooms: row.blockedRooms,
        availableRooms: row.availableRooms,
        status: row.status,
        dailyBreakdown: row.dailyBreakdown
      };
    })
  };

  if (config.includeInternalRows) {
    result.internalBookingRows = bookingRows;
  }

  return result;
}

function normalizeAvailabilityPlannerDays_(value) {
  const allowed = [30, 60, 90, 180, 365, 730];
  const numeric = Math.round(Number(value || 90));
  if (allowed.indexOf(numeric) !== -1) return numeric;
  if (numeric > 0 && numeric <= 730) return numeric;
  return 90;
}

function buildArrivalAllocationPressureMap_(startDate, rangeDays, options) {
  const config = options || {};
  const start = stripTime_(normalizeDateInput_(startDate) || new Date());
  const end = addDays_(start, Math.max(0, Number(rangeDays || 0)));
  const pressureMap = {};

  (config.bookingRows || getSheetObjects_(SHEET_NAMES.BOOKINGS)).forEach(function(row) {
    if (normalizeBookingStatus_(row.status) !== BOOKING_STATUS_CONFIRMED) return;
    const checkIn = normalizeDateInput_(row.check_in);
    if (!checkIn) return;
    const checkInDate = stripTime_(checkIn);
    if (checkInDate.getTime() < start.getTime() || checkInDate.getTime() >= end.getTime()) return;

    const units = Math.max(1, Number(row.qty_rooms || 1));
    const dateKey = formatDateKey_(checkInDate);
    if (!pressureMap[dateKey]) {
      pressureMap[dateKey] = {
        arrivals: 0,
        assignedArrivals: 0,
        unassignedArrivals: 0
      };
    }

    pressureMap[dateKey].arrivals += units;
    if (String(row.room_identifier || '').trim()) {
      pressureMap[dateKey].assignedArrivals += units;
    } else {
      pressureMap[dateKey].unassignedArrivals += units;
    }
  });

  return pressureMap;
}

function buildAvailabilityPlannerPayload_(startDate, rangeDays, overview, options) {
  const config = options || {};
  const roomTypeBreakdown = overview.roomTypeBreakdown || [];
  const arrivalPressureMap = buildArrivalAllocationPressureMap_(startDate, rangeDays, {
    bookingRows: config.bookingRows || overview.internalBookingRows || null
  });
  const plannerColumns = roomTypeBreakdown.map(function(row) {
    return {
      roomTypeId: row.roomTypeId,
      roomTypeName: row.roomTypeName
    };
  });
  const roomTypeBreakdownMap = roomTypeBreakdown.reduce(function(map, row) {
    map[row.roomTypeId] = row;
    return map;
  }, {});

  const plannerRows = (overview.dailyBreakdown || []).map(function(day, index) {
    const pressure = arrivalPressureMap[String(day.date || '').trim()] || {};
    const row = {
      date: day.date,
      totalAvailable: Number(day.availableRooms || 0),
      sold: Number(day.soldRooms || 0),
      blocked: Number(day.blockedRooms || 0),
      arrivals: Number(pressure.arrivals || 0),
      assignedArrivals: Number(pressure.assignedArrivals || 0),
      unassignedArrivals: Number(pressure.unassignedArrivals || 0),
      roomTypes: {}
    };

    plannerColumns.forEach(function(column) {
      const roomType = roomTypeBreakdownMap[column.roomTypeId];
      const roomDay = roomType && roomType.dailyBreakdown ? roomType.dailyBreakdown[index] : null;
      row.roomTypes[column.roomTypeId] = roomDay ? {
        availableRooms: Number(roomDay.availableRooms || 0),
        soldRooms: Number(roomDay.soldRooms || 0),
        blockedRooms: Number(roomDay.blockedRooms || 0),
        totalRooms: Number(roomDay.totalRooms || roomType.totalRooms || 0)
      } : {
        availableRooms: 0,
        soldRooms: 0,
        blockedRooms: 0,
        totalRooms: 0
      };
    });

    return row;
  });

  return {
    ok: true,
    source: overview.source || 'live',
    startDate: formatDateKey_(startDate),
    endDate: formatDateKey_(addDays_(startDate, Math.max(0, rangeDays - 1))),
    days: plannerRows.length,
    totalRooms: Number(overview.totalRooms || 0),
    lowestTotalAvailable: plannerRows.reduce(function(min, row) {
      return Math.min(min, Number(row.totalAvailable || 0));
    }, plannerRows.length ? Number(plannerRows[0].totalAvailable || 0) : 0),
    peakArrivals: plannerRows.reduce(function(max, row) {
      return Math.max(max, Number(row.arrivals || 0));
    }, 0),
    peakUnassignedArrivals: plannerRows.reduce(function(max, row) {
      return Math.max(max, Number(row.unassignedArrivals || 0));
    }, 0),
    peakSold: plannerRows.reduce(function(max, row) {
      return Math.max(max, Number(row.sold || 0));
    }, 0),
    peakBlocked: plannerRows.reduce(function(max, row) {
      return Math.max(max, Number(row.blocked || 0));
    }, 0),
    columns: plannerColumns,
    rows: plannerRows
  };
}

function buildReportsDashboardData_(input) {
  const payload = (input && typeof input === 'object' && Object.prototype.toString.call(input) !== '[object Date]') ? input : {};
  const selectedDate = typeof input === 'string' || Object.prototype.toString.call(input) === '[object Date]'
    ? input
    : (payload.selectedDate || payload.selected_date);
  const asOfDate = normalizeFrontDeskDate_(selectedDate);
  const monthStart = startOfMonth_(asOfDate);
  const yearStart = startOfYear_(asOfDate);
  const lyAsOfDate = shiftDateYears_(asOfDate, -1);
  const lyMonthStart = startOfMonth_(lyAsOfDate);
  const lyYearStart = startOfYear_(lyAsOfDate);
  const bookingNights = getSheetObjects_(SHEET_NAMES.BOOKING_NIGHTS);
  const dailyStats = getSheetObjects_(SHEET_NAMES.DAILY_STATS);
  const bookings = getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const snapshots = getSheetObjects_(SHEET_NAMES.OTB_SNAPSHOTS);
  const currency = getReportingCurrency_();
  const mtd = getPeriodMetrics_(dailyStats, bookingNights, bookings, monthStart, asOfDate);
  const ytd = getPeriodMetrics_(dailyStats, bookingNights, bookings, yearStart, asOfDate);
  const lyMtd = getPeriodMetrics_(dailyStats, bookingNights, bookings, lyMonthStart, lyAsOfDate);
  const lyYtd = getPeriodMetrics_(dailyStats, bookingNights, bookings, lyYearStart, lyAsOfDate);
  const comparisonRange = resolveReportsComparisonRange_(payload, asOfDate);
  const currentMetrics = getPeriodMetrics_(dailyStats, bookingNights, bookings, comparisonRange.current.start, comparisonRange.current.end);
  const comparisonMetrics = getPeriodMetrics_(dailyStats, bookingNights, bookings, comparisonRange.compare.start, comparisonRange.compare.end);
  const otb30 = getPeriodMetrics_(dailyStats, bookingNights, bookings, asOfDate, addDays_(asOfDate, 29));
  const otb90 = getPeriodMetrics_(dailyStats, bookingNights, bookings, asOfDate, addDays_(asOfDate, 89));
  const paceLy = getPaceComparison_(snapshots, asOfDate, 90);
  const pickup7 = getPickupComparison_(snapshots, asOfDate, 7, 90);
  const pickup30 = getPickupComparison_(snapshots, asOfDate, 30, 90);
  const revenueBySource = buildMixShareRows_(summariseNightRowsByField_(bookingNights, monthStart, asOfDate, 'source'));
  const revenueByRoomType = buildMixShareRows_(summariseNightRowsByField_(bookingNights, monthStart, asOfDate, 'room_type_name'));
  const insights = buildPortfolioInsights_(bookings, monthStart, asOfDate);

  return {
    ok: true,
    selectedDate: formatDateKey_(asOfDate),
    currency: currency,
    summary: {
      mtdRevenue: roundCurrency_(mtd.roomRevenue),
      ytdRevenue: roundCurrency_(ytd.roomRevenue),
      adr: roundCurrency_(mtd.adr),
      occupancyPct: mtd.occupancyPct,
      revpar: roundCurrency_(mtd.revpar),
      roomNights: mtd.roomNights,
      availableRoomNights: mtd.availableRoomNights,
      alos: roundCurrency_(mtd.alos),
      directWebsiteSharePct: insights.directWebsiteSharePct,
      cancellationsMtd: insights.cancellationsCount,
      averageLeadTimeDays: insights.averageLeadTimeDays
    },
    summaryRows: [
      buildReportsMetricRow_('MTD', mtd),
      buildReportsMetricRow_('MTD LY', lyMtd),
      buildReportsMetricRow_('YTD', ytd),
      buildReportsMetricRow_('YTD LY', lyYtd)
    ],
    comparison: {
      currentLabel: comparisonRange.current.label,
      compareLabel: comparisonRange.compare.label,
      mode: comparisonRange.mode,
      currentMetrics: buildReportsMetricRow_(comparisonRange.current.label, currentMetrics),
      compareMetrics: buildReportsMetricRow_(comparisonRange.compare.label, comparisonMetrics)
    },
    otb: {
      next30: buildReportsMetricRow_('OTB Next 30', otb30),
      next90: buildReportsMetricRow_('OTB Next 90', otb90),
      pace: {
        available: paceLy.available,
        currentRevenue: roundCurrency_(paceLy.currentRevenue),
        lastYearRevenue: roundCurrency_(paceLy.lastYearRevenue),
        currentRoomNights: paceLy.currentRoomNights,
        lastYearRoomNights: paceLy.lastYearRoomNights
      },
      pickup7: {
        available: pickup7.available,
        currentRevenue: roundCurrency_(pickup7.currentRevenue),
        previousRevenue: roundCurrency_(pickup7.previousRevenue),
        currentRoomNights: pickup7.currentRoomNights,
        previousRoomNights: pickup7.previousRoomNights
      },
      pickup30: {
        available: pickup30.available,
        currentRevenue: roundCurrency_(pickup30.currentRevenue),
        previousRevenue: roundCurrency_(pickup30.previousRevenue),
        currentRoomNights: pickup30.currentRoomNights,
        previousRoomNights: pickup30.previousRoomNights
      }
    },
    revenueMix: {
      bySource: revenueBySource,
      byRoomType: revenueByRoomType
    }
  };
}

function buildReportsMetricRow_(label, metrics) {
  return {
    label: label,
    roomRevenue: roundCurrency_(metrics.roomRevenue),
    roomNights: Number(metrics.roomNights || 0),
    occupancyPct: Number(metrics.occupancyPct || 0),
    adr: roundCurrency_(metrics.adr),
    revpar: roundCurrency_(metrics.revpar),
    alos: roundCurrency_(metrics.alos)
  };
}

function buildMixShareRows_(rows) {
  const validRows = (rows || []).filter(function(row) {
    return row && row[0] && row[0] !== 'No data yet';
  }).map(function(row) {
    return {
      label: row[0],
      roomRevenue: Number(row[1] || 0),
      roomNights: Number(row[2] || 0)
    };
  });
  const totalRevenue = validRows.reduce(function(sum, row) {
    return sum + Number(row.roomRevenue || 0);
  }, 0);

  return validRows.length ? validRows.map(function(row) {
    return {
      label: row.label,
      roomRevenue: roundCurrency_(row.roomRevenue),
      roomNights: row.roomNights,
      revenueSharePct: safeDivide_(row.roomRevenue, totalRevenue)
    };
  }) : [];
}

function resolveReportsComparisonRange_(payload, asOfDate) {
  let currentStart = normalizeDateInput_(payload.current_start || payload.currentStart) || startOfMonth_(asOfDate);
  let currentEnd = normalizeDateInput_(payload.current_end || payload.currentEnd) || asOfDate;
  if (currentEnd.getTime() < currentStart.getTime()) {
    currentEnd = currentStart;
  }
  const currentDays = Math.max(1, Math.round((stripTime_(currentEnd).getTime() - stripTime_(currentStart).getTime()) / 86400000) + 1);
  const mode = String(payload.compare_mode || payload.compareMode || 'previous').trim().toLowerCase();
  let compareStart;
  let compareEnd;
  let compareLabel;

  if (mode === 'custom') {
    compareStart = normalizeDateInput_(payload.compare_start || payload.compareStart) || addDays_(currentStart, -currentDays);
    compareEnd = normalizeDateInput_(payload.compare_end || payload.compareEnd) || addDays_(compareStart, currentDays - 1);
    if (compareEnd.getTime() < compareStart.getTime()) {
      compareEnd = compareStart;
    }
    compareLabel = 'Custom Comparison';
  } else if (mode === 'ly' || mode === 'same-ly' || mode === 'last-year') {
    compareStart = shiftDateYears_(currentStart, -1);
    compareEnd = shiftDateYears_(currentEnd, -1);
    compareLabel = 'Same Period LY';
  } else {
    compareEnd = addDays_(currentStart, -1);
    compareStart = addDays_(compareEnd, -(currentDays - 1));
    compareLabel = 'Previous Period';
  }

  return {
    mode: mode,
    current: {
      start: currentStart,
      end: currentEnd,
      label: formatDateKey_(currentStart) + ' to ' + formatDateKey_(currentEnd)
    },
    compare: {
      start: compareStart,
      end: compareEnd,
      label: compareLabel + ' (' + formatDateKey_(compareStart) + ' to ' + formatDateKey_(compareEnd) + ')'
    }
  };
}

function mapBookingRowToAdminDisplay_(row) {
  const context = arguments[1] || {};
  const reportingCurrency = context.reportingCurrency || getReportingCurrency_();
  const defaultBookingCurrency = context.defaultBookingCurrency || getDefaultBookingCurrency_();
  const bookingValue = getOperationalBookingValueGbp_(row, {
    defaultBookingCurrency: defaultBookingCurrency
  });
  const bookingValueOriginal = getStoredBookingOriginalValue_(row);
  const bookingCurrency = getStoredBookingCurrency_(row, defaultBookingCurrency);
  const fxRateToGbp = roundFxRate_(row.fx_rate_to_gbp || (bookingCurrency === reportingCurrency ? 1 : 0));
  const pricingSource = String(row.pricing_source || '').trim();
  const pricingReferenceId = String(row.pricing_reference_id || '').trim();
  const amountPaid = roundCurrency_(row.amount_paid || 0);
  const taxAmount = roundCurrency_(row.tax_amount || 0);
  const checkIn = normalizeDateInput_(row.check_in);
  const checkOut = normalizeDateInput_(row.check_out);
  const history = context.historyIndex ? context.historyIndex[String(row.booking_id || '').trim()] : null;
  const paymentMethod = normalizePaymentMethod_(row.payment_method || '');
  const paymentNotes = String(row.payment_notes || '').trim();
  const guestPreferences = String(row.guest_preferences || '').trim();
  const rawGuestNote = String(row.notes || '').trim();
  const rawInternalNote = String(row.internal_notes || '').trim();
  const balanceDue = calculateBalanceDue_(bookingValue, amountPaid);
  const roomIdentifier = String(row.room_identifier || '').trim();
  const roomIndex = context.roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const matchedRoom = resolveAssignedRoomFromIndex_(roomIndex, roomIdentifier, String(row.room_type_id || '').trim());
  const roomAllowedSetups = matchedRoom ? parseAllowedSetups_(matchedRoom.allowedSetups || matchedRoom.allowed_setups) : [];
  const normalizedBedSetup = normalizeBedSetup_(row.bed_setup || '');
  const normalizedSource = normalizeBookingSource_(row.source || '');
  const guestNote = normalizedSource === 'Direct Website' ? rawGuestNote : '';
  const guestNoteLabel = guestNote ? 'Guest / Website Note' : '';
  const internalNotes = rawInternalNote || (!guestNote ? rawGuestNote : '');
  const hasBalanceDue = balanceDue > 0.009;
  const hasGuestNote = !!guestNote;
  const hasInternalNotes = !!internalNotes;
  const hasAttentionNotes = hasGuestNote || hasInternalNotes;
  const hasPreferences = !!guestPreferences;
  const hasBedSetupMismatch = matchedRoom ? !isRequestedSetupCompatibleWithRoom_(normalizedBedSetup, matchedRoom) : false;
  const prepStatusKey = getArrivalPrepStatusKey_({
    roomIdentifier: roomIdentifier,
    hasBalanceDue: hasBalanceDue,
    hasAttentionNotes: hasAttentionNotes,
    hasPreferences: hasPreferences,
    hasBedSetupMismatch: hasBedSetupMismatch
  });
  const prepSignals = [];
  if (!roomIdentifier) prepSignals.push('Room not assigned');
  if (hasBedSetupMismatch) prepSignals.push('Bed setup mismatch');
  if (hasGuestNote) prepSignals.push('Guest note');
  if (hasInternalNotes) prepSignals.push('Internal note');
  if (hasPreferences) prepSignals.push('Preference recorded');
  if (hasBalanceDue) prepSignals.push('Balance due');

  return {
    bookingId: String(row.booking_id || '').trim(),
    requestId: String(row.request_id || '').trim(),
    guestName: String(row.guest_name || '').trim(),
    guestPhone: String(row.guest_phone || '').trim(),
    guestEmail: String(row.guest_email || '').trim(),
    country: String(row.country || '').trim(),
    checkIn: checkIn ? formatDateKey_(checkIn) : '',
    checkInTime: normalizeTimeInput_(row.check_in_time),
    checkOut: checkOut ? formatDateKey_(checkOut) : '',
    checkOutTime: normalizeTimeInput_(row.check_out_time),
    roomTypeId: String(row.room_type_id || '').trim(),
    roomTypeName: String(row.room_type_name || row.room_type_id || '').trim(),
    roomIdentifier: roomIdentifier,
    isRoomAssigned: !!roomIdentifier,
    isKnownRoomAssignment: !!matchedRoom,
    roomDisplayName: matchedRoom ? matchedRoom.roomName : roomIdentifier,
    roomCode: matchedRoom ? matchedRoom.roomCode : roomIdentifier,
    roomShortLabel: buildShortRoomLabel_(String(row.room_type_id || '').trim(), matchedRoom ? matchedRoom.roomCode : roomIdentifier, roomIdentifier),
    roomTypeShortLabel: getRoomTypeShortLabel_(String(row.room_type_id || '').trim(), String(row.room_type_name || row.room_type_id || '').trim()),
    roomDefaultSetup: matchedRoom ? matchedRoom.defaultSetup : '',
    roomAllowedSetups: roomAllowedSetups,
    bedSetup: normalizedBedSetup,
    isBedSetupCompatible: matchedRoom ? isRequestedSetupCompatibleWithRoom_(normalizedBedSetup, matchedRoom) : true,
    source: normalizedSource,
    status: normalizeBookingStatus_(row.status),
    paymentStatus: normalizePaymentStatus_(row.payment_status, bookingValue, amountPaid),
    paymentMethod: paymentMethod,
    paymentNotes: paymentNotes,
    bookingValue: bookingValue,
    bookingValueOriginal: bookingValueOriginal,
    bookingValueGbp: bookingValue,
    bookingCurrency: bookingCurrency,
    reportingCurrency: reportingCurrency,
    fxRateToGbp: fxRateToGbp,
    pricingSource: pricingSource,
    pricingReferenceId: pricingReferenceId,
    taxAmount: taxAmount,
    amountPaid: amountPaid,
    balanceDue: balanceDue,
    qtyRooms: Math.max(1, Number(row.qty_rooms || 1)),
    adults: Math.max(0, Number(row.adults || 0)),
    children: Math.max(0, Number(row.children || 0)),
    guests: Math.max(0, Number(row.guests || 0)),
    guestPreferences: guestPreferences,
    guestNote: guestNote,
    guestNoteLabel: guestNoteLabel,
    internalNotes: internalNotes,
    hasGuestNote: hasGuestNote,
    hasInternalNotes: hasInternalNotes,
    hasAttentionNotes: hasAttentionNotes,
    hasPreferences: hasPreferences,
    hasBalanceDue: hasBalanceDue,
    hasBedSetupMismatch: hasBedSetupMismatch,
    prepStatusKey: prepStatusKey,
    prepStatusLabel: getArrivalPrepStatusLabel_(prepStatusKey),
    prepPriority: getArrivalPrepPriority_(prepStatusKey),
    prepSignals: prepSignals,
    isRepeatGuest: history ? history.isRepeatGuest : false,
    repeatStayCount: history ? history.repeatStayCount : 0,
    lastStayCheckOut: history ? history.lastStayCheckOut : '',
    guestHistory: history ? history.previousStays : [],
    lastPreference: history ? history.lastPreference : guestPreferences,
    prepSummary: buildBookingPrepSummary_(roomIdentifier, balanceDue, guestPreferences, paymentNotes, guestNote || internalNotes),
    notes: rawGuestNote
  };
}

function normalizeFrontDeskDate_(selectedDate) {
  return normalizeDateInput_(selectedDate) || stripTime_(new Date());
}

// -----------------------------------------------------------------------------
// Core logic helpers
// -----------------------------------------------------------------------------

function normalizeCommercialRuleType_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  return COMMERCIAL_RULE_TYPE_VALUES.indexOf(normalized) !== -1 ? normalized : 'seasonal';
}

function formatCommercialRuleTypeLabel_(ruleType) {
  return normalizeCommercialRuleType_(ruleType) === 'special' ? 'Special override' : 'Seasonal override';
}

function normalizeDirectDiscountType_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized === 'percent' || normalized === 'percentage' || normalized === '%') return 'Percent';
  if (normalized === 'fixed' || normalized === 'amount' || normalized === 'flat') return 'Fixed';
  return 'None';
}

function normalizeOptionalPositiveCurrency_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  const amount = roundCurrency_(raw);
  return isFinite(Number(amount)) && Number(amount) > 0 ? amount : '';
}

function normalizeDirectBookingDiscountFields_(source) {
  const row = source || {};
  const referencePrice = normalizeOptionalPositiveCurrency_(
    row.public_reference_price != null ? row.public_reference_price :
      (row.publicReferencePrice != null ? row.publicReferencePrice :
        (row.ota_reference_price != null ? row.ota_reference_price :
          (row.otaReferencePrice != null ? row.otaReferencePrice :
            (row.comparison_price != null ? row.comparison_price : row.comparisonPrice))))
  );
  const discountType = normalizeDirectDiscountType_(row.direct_discount_type != null ? row.direct_discount_type : row.directDiscountType);
  const rawValue = String(row.direct_discount_value != null ? row.direct_discount_value : (row.directDiscountValue != null ? row.directDiscountValue : '')).trim();
  let discountValue = '';
  if (discountType === 'Percent' && rawValue !== '') {
    const percent = Number(rawValue);
    discountValue = isFinite(percent) && percent > 0 && percent < 100 ? Math.round(percent * 100) / 100 : '';
  } else if (discountType === 'Fixed') {
    discountValue = normalizeOptionalPositiveCurrency_(rawValue);
  }

  return {
    publicReferencePrice: referencePrice,
    directDiscountType: discountType,
    directDiscountValue: discountValue
  };
}

function hasDirectBookingDiscountValue_(fields) {
  const config = fields || {};
  return normalizeDirectDiscountType_(config.directDiscountType) !== 'None' &&
    Number(config.directDiscountValue || 0) > 0;
}

function getDirectBookingReferencePrice_(fields, fallbackReferencePrice) {
  const config = fields || {};
  const explicitReferencePrice = Number(config.publicReferencePrice || 0);
  if (explicitReferencePrice > 0) return explicitReferencePrice;
  const fallbackPrice = Number(fallbackReferencePrice || 0);
  return fallbackPrice > 0 ? fallbackPrice : 0;
}

function hasDirectBookingDiscountConfig_(fields, fallbackReferencePrice) {
  return hasDirectBookingDiscountValue_(fields) &&
    getDirectBookingReferencePrice_(fields, fallbackReferencePrice) > 0;
}

function validateDirectBookingDiscountFields_(fields, fallbackReferencePrice) {
  const config = fields || {};
  if (config.publicReferencePrice !== '' && (!isFinite(Number(config.publicReferencePrice)) || Number(config.publicReferencePrice) <= 0)) {
    throw new Error('Public reference price must be a positive amount.');
  }
  if (normalizeDirectDiscountType_(config.directDiscountType) !== 'None') {
    if (config.directDiscountValue === '' || Number(config.directDiscountValue || 0) <= 0) {
      throw new Error('Direct booking discount value must be greater than zero.');
    }
    const referencePrice = getDirectBookingReferencePrice_(config, fallbackReferencePrice);
    if (!(referencePrice > 0)) {
      throw new Error('Add a public reference price before setting a direct booking discount.');
    }
    if (normalizeDirectDiscountType_(config.directDiscountType) === 'Fixed' && Number(config.directDiscountValue || 0) >= referencePrice) {
      throw new Error('Fixed direct booking discount must be lower than the reference price.');
    }
  }
}

function formatPercentValue_(value) {
  return Utilities.formatString('%.1f%%', Number(value || 0) * 100);
}

function isYesLike_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  return normalized === 'yes' || normalized === 'true' || normalized === 'active' || normalized === '1' || normalized === 'on';
}

function toDateTimeMs_(value) {
  if (!value) return 0;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.getTime();
  }
  const parsed = new Date(value);
  const parsedMs = parsed.getTime();
  return isFinite(parsedMs) ? parsedMs : 0;
}

function normalizeBaseRateRow_(row, rowNumber, roomTypeNameMap) {
  const rawRoomTypeId = String(row.room_type_id || '').trim();
  const roomTypeId = rawRoomTypeId ? resolveRoomTypeId_(rawRoomTypeId) : '';
  const extraGuestFeeRaw = String(row.extra_guest_fee == null ? '' : row.extra_guest_fee).trim();
  const directDiscount = normalizeDirectBookingDiscountFields_(row);
  return {
    rowNumber: Number(rowNumber || 0),
    roomTypeId: roomTypeId,
    roomTypeName: String(row.room_type_name || (roomTypeNameMap && roomTypeNameMap[roomTypeId]) || (roomTypeId ? getRoomTypeNameById_(roomTypeId) : '') || '').trim(),
    baseRate: roundCurrency_(row.base_rate || 0),
    extraGuestFee: extraGuestFeeRaw === '' ? '' : roundCurrency_(row.extra_guest_fee || 0),
    hasExtraGuestFee: extraGuestFeeRaw !== '',
    publicReferencePrice: directDiscount.publicReferencePrice,
    directDiscountType: directDiscount.directDiscountType,
    directDiscountValue: directDiscount.directDiscountValue,
    active: isYesLike_(row.active || 'Yes'),
    updatedAt: row.updated_at || '',
    updatedAtMs: toDateTimeMs_(row.updated_at)
  };
}

function getBaseRateRows_(options) {
  const config = options || {};
  const targetRoomTypeId = String(config.roomTypeId || '').trim();
  return getSheetObjects_(SHEET_NAMES.BASE_RATES)
    .map(function(row, index) {
      return normalizeBaseRateRow_(row, index + 2, config.roomTypeNameMap);
    })
    .filter(function(row) {
      if (!row.roomTypeId) return false;
      if (config.activeOnly && !row.active) return false;
      if (targetRoomTypeId && row.roomTypeId !== targetRoomTypeId) return false;
      return true;
    });
}

function normalizeCommercialControlRow_(row, rowNumber, roomTypeNameMap) {
  const roomTypeId = String(row.room_type_id || '').trim();
  const startDate = normalizeDateInput_(row.start_date);
  const endDate = normalizeDateInput_(row.end_date);
  const overrideRaw = String(row.override_price == null ? '' : row.override_price).trim();
  const overridePrice = overrideRaw === '' ? '' : roundCurrency_(row.override_price);
  const overbookingRaw = String(row.overbooking_allowance == null ? '' : row.overbooking_allowance).trim();
  const overbookingAllowance = Math.max(0, Number(overbookingRaw === '' ? 0 : row.overbooking_allowance || 0));
  const directDiscount = normalizeDirectBookingDiscountFields_(row);

  return {
    rowNumber: Number(rowNumber || 0),
    controlId: String(row.control_id || '').trim(),
    roomTypeId: roomTypeId,
    roomTypeName: String(row.room_type_name || (roomTypeNameMap && roomTypeNameMap[roomTypeId]) || (roomTypeId ? getRoomTypeNameById_(roomTypeId) : '') || '').trim(),
    ruleType: normalizeCommercialRuleType_(row.rule_type),
    startDate: startDate,
    endDate: endDate,
    overridePrice: overridePrice,
    hasOverridePrice: overrideRaw !== '',
    publicReferencePrice: directDiscount.publicReferencePrice,
    directDiscountType: directDiscount.directDiscountType,
    directDiscountValue: directDiscount.directDiscountValue,
    overbookingAllowance: overbookingAllowance,
    active: isYesLike_(row.active || 'Yes'),
    note: String(row.note || '').trim(),
    createdAt: row.created_at || '',
    updatedAt: row.updated_at || row.created_at || '',
    updatedAtMs: Math.max(toDateTimeMs_(row.updated_at), toDateTimeMs_(row.created_at))
  };
}

function getCommercialControlRows_(options) {
  const config = options || {};
  const targetRoomTypeId = String(config.roomTypeId || '').trim();
  return getSheetObjects_(SHEET_NAMES.COMMERCIAL_CONTROLS)
    .map(function(row, index) {
      return normalizeCommercialControlRow_(row, index + 2, config.roomTypeNameMap);
    })
    .filter(function(row) {
      if (!row.controlId || !row.roomTypeId || !row.startDate || !row.endDate) return false;
      if (config.activeOnly && !row.active) return false;
      if (targetRoomTypeId && row.roomTypeId !== targetRoomTypeId) return false;
      return true;
    });
}

function commercialControlAppliesToDate_(row, date) {
  if (!row || !row.startDate || !row.endDate) return false;
  const targetTime = stripTime_(date).getTime();
  return targetTime >= stripTime_(row.startDate).getTime() && targetTime <= stripTime_(row.endDate).getTime();
}

function getCommercialRulePriority_(ruleType) {
  return normalizeCommercialRuleType_(ruleType) === 'special' ? 2 : 1;
}

function compareCommercialControlPriority_(a, b) {
  const priorityDelta = getCommercialRulePriority_(b && b.ruleType) - getCommercialRulePriority_(a && a.ruleType);
  if (priorityDelta !== 0) return priorityDelta;
  const updatedDelta = Number(b && b.updatedAtMs || 0) - Number(a && a.updatedAtMs || 0);
  if (updatedDelta !== 0) return updatedDelta;
  return Number(b && b.rowNumber || 0) - Number(a && a.rowNumber || 0);
}

function selectCommercialControlForDate_(roomTypeId, date, prefetchedRows, predicate) {
  const targetRoomTypeId = String(roomTypeId || '').trim();
  let selected = null;

  (prefetchedRows || getCommercialControlRows_({ activeOnly: true, roomTypeId: targetRoomTypeId })).forEach(function(row) {
    if (!row.active || row.roomTypeId !== targetRoomTypeId || !commercialControlAppliesToDate_(row, date)) return;
    if (predicate && !predicate(row)) return;
    if (!selected || compareCommercialControlPriority_(row, selected) < 0) {
      selected = row;
    }
  });

  return selected;
}

function getCommercialPriceControlForDate_(roomTypeId, date, prefetchedRows) {
  return selectCommercialControlForDate_(roomTypeId, date, prefetchedRows, function(row) {
    return row.hasOverridePrice || hasDirectBookingDiscountConfig_(row);
  });
}

function getCommercialOverbookingControlForDate_(roomTypeId, date, prefetchedRows) {
  return selectCommercialControlForDate_(roomTypeId, date, prefetchedRows, function(row) {
    return Number(row.overbookingAllowance || 0) > 0;
  });
}

function getCommercialControlForDate_(roomTypeId, date, prefetchedRows) {
  return selectCommercialControlForDate_(roomTypeId, date, prefetchedRows, null);
}

function getSellableInventoryTotalForDate_(date, roomTypeId, prefetchedControls, physicalInventory) {
  const baseInventory = Number(physicalInventory != null ? physicalInventory : getRoomInventory_(roomTypeId));
  const control = getCommercialOverbookingControlForDate_(roomTypeId, date, prefetchedControls);
  return baseInventory + Number(control && control.overbookingAllowance || 0);
}

function getLegacyFallbackBaseRateForRoomType_(roomTypeId) {
  const rows = getOpenRateRowsForType_(roomTypeId);
  if (!rows.length) return 0;
  let selected = rows[0];
  rows.forEach(function(row) {
    const currentDate = normalizeDateInput_(row.date);
    const selectedDate = normalizeDateInput_(selected.date);
    if (!selectedDate || (currentDate && currentDate.getTime() > selectedDate.getTime())) {
      selected = row;
    }
  });
  return roundCurrency_(selected.base_rate || 0);
}

function buildCommercialControlsData_(selectedDate, options) {
  const config = options || {};
  const anchorDate = normalizeFrontDeskDate_(selectedDate);
  const roomTypeNameMap = config.roomTypeNameMap || buildRoomTypeNameMap_(getSheetObjects_(SHEET_NAMES.ROOM_TYPES));
  const controls = getCommercialControlRows_({
    activeOnly: false,
    roomTypeNameMap: roomTypeNameMap
  }).sort(function(a, b) {
    const liveDelta = Number(b.active && commercialControlAppliesToDate_(b, anchorDate)) - Number(a.active && commercialControlAppliesToDate_(a, anchorDate));
    if (liveDelta !== 0) return liveDelta;
    const activeDelta = Number(b.active) - Number(a.active);
    if (activeDelta !== 0) return activeDelta;
    const aStart = a.startDate ? stripTime_(a.startDate).getTime() : 0;
    const bStart = b.startDate ? stripTime_(b.startDate).getTime() : 0;
    if (aStart !== bStart) return aStart - bStart;
    return b.updatedAtMs - a.updatedAtMs;
  });

  return {
    ok: true,
    selectedDate: formatDateKey_(anchorDate),
    summary: {
      totalRules: controls.length,
      activeRules: controls.filter(function(row) { return row.active; }).length,
      liveRules: controls.filter(function(row) { return row.active && commercialControlAppliesToDate_(row, anchorDate); }).length
    },
    controls: controls.map(function(row) {
      return {
        controlId: row.controlId,
        roomTypeId: row.roomTypeId,
        roomTypeName: row.roomTypeName,
        ruleType: row.ruleType,
        startDate: formatDateKey_(row.startDate),
        endDate: formatDateKey_(row.endDate),
        overridePrice: row.hasOverridePrice ? row.overridePrice : '',
        publicReferencePrice: row.publicReferencePrice,
        directDiscountType: row.directDiscountType,
        directDiscountValue: row.directDiscountValue,
        overbookingAllowance: row.overbookingAllowance,
        active: row.active,
        note: row.note,
        updatedAtLabel: row.updatedAt ? Utilities.formatDate(new Date(row.updatedAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm') : '',
        isLive: row.active && commercialControlAppliesToDate_(row, anchorDate)
      };
    })
  };
}

function buildRateBoardData_(selectedDate) {
  const targetDate = normalizeFrontDeskDate_(selectedDate);
  const currency = getDefaultBookingCurrency_();
  const roomTypes = getRateBoardRoomTypes_();
  const baseStartedAt = Date.now();
  const baseRates = buildBaseRatesData_();
  logTiming_('rateBoard:readBaseRates', baseStartedAt);
  const baseRateMap = {};
  (baseRates.rates || []).forEach(function(rate) {
    baseRateMap[rate.roomTypeId] = rate;
  });
  const ladder = getRateLadderConfig_();
  const leadBaseRate = Number((baseRateMap[ladder.leadRoomTypeId] && baseRateMap[ladder.leadRoomTypeId].baseRate) || 0);
  const controlsStartedAt = Date.now();
  const commercialControls = getCommercialControlRows_({ activeOnly: true });
  const baseRateRows = getBaseRateRows_({ activeOnly: true });
  const legacyRateRows = getSheetObjects_(SHEET_NAMES.RATES).filter(function(row) {
    return String(row.status || '').trim().toLowerCase() === 'open';
  });
  logTiming_('rateBoard:readCommercialControls', controlsStartedAt);
  const availabilityStartedAt = Date.now();
  const nextDate = addDays_(targetDate, 1);
  const bookingCountMap = buildConfirmedBookingRoomCountMap_(getSheetObjects_(SHEET_NAMES.BOOKINGS), {
    startDate: targetDate,
    endDateExclusive: nextDate
  }).byDateRoom;
  const blockedCountMap = buildBlockedDateQtyMap_(getSheetObjects_(SHEET_NAMES.BLOCKED_DATES), {
    startDate: targetDate,
    endDateExclusive: nextDate
  }).byDateRoom;
  logTiming_('rateBoard:readAvailabilityCheap', availabilityStartedAt);
  const otaRows = getSheetObjects_(SHEET_NAMES.OTA_UPDATE_WORKFLOW).map(function(row, index) {
    return normalizeOtaWorkflowRow_(row, index + 2);
  });

  const rowsStartedAt = Date.now();
  const rows = roomTypes.map(function(roomType) {
    const rateRow = baseRateMap[roomType.roomTypeId] || {};
    const differential = Number(ladder.differentials[roomType.roomTypeId] || 0);
    const expectedBaseRate = roundCurrency_(leadBaseRate + differential);
    const currentBaseRate = Number(rateRow.baseRate || 0);
    const extraGuestFee = rateRow.extraGuestFee === '' || rateRow.extraGuestFee == null ? '' : Number(rateRow.extraGuestFee || 0);
    const guests = Number(INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE[roomType.roomTypeId] || roomType.maxGuests || 2);
    const roomLegacyRateRows = legacyRateRows.filter(function(row) {
      return String(row.room_type_id || '').trim() === roomType.roomTypeId;
    });
    const liveCommercial = resolveLiveCommercialDate_(targetDate, roomType.roomTypeId, guests, {
      commercialControls: commercialControls,
      baseRateRows: baseRateRows,
      legacyRateRows: roomLegacyRateRows,
      physicalInventory: Number(roomType.inventoryTotal || 0),
      bookingCountMap: bookingCountMap,
      blockedCountMap: blockedCountMap
    });
    const priceControl = getCommercialPriceControlForDate_(roomType.roomTypeId, targetDate, commercialControls);
    const overbookingControl = getCommercialOverbookingControlForDate_(roomType.roomTypeId, targetDate, commercialControls);
    const otaState = getLatestOtaWorkflowStateForDate_(otaRows, targetDate, roomType.roomTypeId);
    const occupancyPct = safeDivide_(Number(liveCommercial && liveCommercial.soldRooms || 0), Number(liveCommercial && liveCommercial.sellableInventory || 0));
    const pressure = buildRateBoardPressureState_({
      occupancyPct: occupancyPct,
      exactDateMetrics: {
        availableRooms: Number(liveCommercial && liveCommercial.availableRooms || 0),
        soldRooms: Number(liveCommercial && liveCommercial.soldRooms || 0),
        sellableInventory: Number(liveCommercial && liveCommercial.sellableInventory || 0)
      }
    });
    const followsLadder = roomType.roomTypeId === ladder.leadRoomTypeId
      ? true
      : Math.abs(currentBaseRate - expectedBaseRate) < 0.01;
    const overrideState = buildRateBoardOverrideState_(roomType.roomTypeId, ladder.leadRoomTypeId, followsLadder, priceControl);

    return {
      date: formatDateKey_(targetDate),
      roomTypeId: roomType.roomTypeId,
      roomTypeName: roomType.roomTypeName,
      guests: guests,
      inventoryTotal: Number(roomType.inventoryTotal || 0),
      baseRate: currentBaseRate,
      extraGuestFee: extraGuestFee,
      differential: differential,
      expectedBaseRate: expectedBaseRate,
      followsLadder: followsLadder,
      liveSellPrice: liveCommercial && liveCommercial.finalNightlyRate != null ? Number(liveCommercial.finalNightlyRate) : null,
      livePricingSource: String(liveCommercial && liveCommercial.pricingSource || '').trim() || 'base_rate',
      livePricingReferenceId: String(liveCommercial && liveCommercial.pricingReferenceId || '').trim(),
      overbookingAllowance: Number(overbookingControl && overbookingControl.overbookingAllowance || 0),
      availableRooms: Number(liveCommercial && liveCommercial.availableRooms || 0),
      sellableInventory: Number(liveCommercial && liveCommercial.sellableInventory || 0),
      soldRooms: Number(liveCommercial && liveCommercial.soldRooms || 0),
      occupancyPct: occupancyPct,
      demandBand: 'Not loaded',
      recommendationAction: 'Signals not loaded',
      recommendationConfidence: 'Not loaded',
      signalStatus: 'not_loaded',
      signalMeta: 'Commercial signals not loaded yet. Use Refresh Signals if needed.',
      pressureTone: pressure.tone,
      pressureLabel: pressure.label,
      pressureMeta: pressure.meta,
      pressureRank: pressure.rank,
      overrideTone: overrideState.tone,
      overrideLabel: overrideState.label,
      overrideMeta: overrideState.meta,
      hasDateOverride: overrideState.hasDateOverride,
      hasLadderBreakaway: overrideState.hasLadderBreakaway,
      overrideStates: overrideState.states,
      otaState: otaState.status,
      otaMeta: otaState.meta,
      otaUpdatedAtLabel: otaState.updatedAtLabel
    };
  });
  logTiming_('rateBoard:buildRows', rowsStartedAt);

  const result = {
    ok: true,
    selectedDate: formatDateKey_(targetDate),
    currency: currency,
    signalStatus: 'not_loaded',
    signalMessage: 'Commercial signals not loaded yet. Use Refresh Signals if needed.',
    ladder: {
      leadRoomTypeId: ladder.leadRoomTypeId,
      leadRoomTypeName: ladder.leadRoomTypeName,
      leadBaseRate: leadBaseRate,
      updatedAtLabel: ladder.updatedAtLabel,
      lastAppliedAtLabel: ladder.lastAppliedAtLabel,
      lastReason: ladder.lastReason,
      differentials: ladder.differentials
    },
    summary: {
      liveOverrideCount: rows.filter(function(row) { return row.hasDateOverride; }).length,
      baseOverrideCount: rows.filter(function(row) { return row.hasLadderBreakaway; }).length,
      highPressureCount: rows.filter(function(row) { return row.pressureTone === 'red'; }).length,
      watchPressureCount: rows.filter(function(row) { return row.pressureTone === 'amber'; }).length,
      pendingOtaCount: rows.filter(function(row) { return String(row.otaState || '').trim() !== 'Updated' && String(row.otaState || '').trim() !== 'Not relevant'; }).length,
      updatedOtaCount: rows.filter(function(row) { return String(row.otaState || '').trim() === 'Updated'; }).length
    },
    rows: rows
  };
  return result;
}

function validateRateCalendarInput_(input) {
  const payload = input || {};
  const rawMode = String(payload.mode || payload.view || '').trim().toLowerCase();
  const hasDayRangeInput = payload.days != null || payload.range_days != null || payload.rangeDays != null;
  const mode = rawMode === 'days' || rawMode.indexOf('days_') === 0 || (hasDayRangeInput && rawMode !== 'month')
    ? 'days'
    : 'month';
  const currentMonthStart = getRateCalendarCurrentMonthStart_();
  const maxMonthStart = addMonthsStart_(currentMonthStart, 12);

  if (mode === 'month') {
    const requestedMonth = parseRateCalendarMonthInput_(payload.month || payload.selected_month || payload.selectedMonth);
    const monthStart = clampRateCalendarMonth_(requestedMonth || currentMonthStart, currentMonthStart, maxMonthStart);
    return {
      mode: 'month',
      startDate: monthStart,
      days: getDaysInMonth_(monthStart),
      month: formatRateCalendarMonthKey_(monthStart),
      minMonth: formatRateCalendarMonthKey_(currentMonthStart),
      maxMonth: formatRateCalendarMonthKey_(maxMonthStart)
    };
  }

  const startDate = normalizeDateInput_(payload.start_date || payload.startDate || payload.date) || stripTime_(new Date());
  const viewDaysMatch = rawMode.match(/^days_(\d+)$/);
  const requestedDays = Number(payload.days || payload.range_days || payload.rangeDays || (viewDaysMatch ? viewDaysMatch[1] : 14));
  return {
    mode: 'days',
    startDate: startDate,
    days: requestedDays === 30 ? 30 : 14,
    month: '',
    minMonth: formatRateCalendarMonthKey_(currentMonthStart),
    maxMonth: formatRateCalendarMonthKey_(maxMonthStart)
  };
}

function getRateCalendarCurrentMonthStart_() {
  const today = stripTime_(new Date());
  return new Date(today.getFullYear(), today.getMonth(), 1);
}

function parseRateCalendarMonthInput_(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const monthIndex = Number(match[2]) - 1;
  if (!isFinite(year) || monthIndex < 0 || monthIndex > 11) return null;
  return new Date(year, monthIndex, 1);
}

function addMonthsStart_(date, months) {
  const base = date || getRateCalendarCurrentMonthStart_();
  return new Date(base.getFullYear(), base.getMonth() + Number(months || 0), 1);
}

function clampRateCalendarMonth_(monthStart, minMonthStart, maxMonthStart) {
  const target = monthStart || minMonthStart;
  if (target.getTime() < minMonthStart.getTime()) return minMonthStart;
  if (target.getTime() > maxMonthStart.getTime()) return maxMonthStart;
  return target;
}

function getDaysInMonth_(monthStart) {
  return new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0).getDate();
}

function formatRateCalendarMonthKey_(monthStart) {
  return String(monthStart.getFullYear()) + '-' + String(monthStart.getMonth() + 1).padStart(2, '0');
}

function validateRateCalendarCellInput_(input) {
  const payload = input || {};
  const date = normalizeDateInput_(payload.date || payload.selected_date || payload.selectedDate);
  const roomTypeId = resolveRoomTypeId_(payload.room_type_id || payload.roomTypeId || payload.room_type || payload.roomType || '');
  const overridePrice = roundCurrency_(payload.override_price || payload.overridePrice || payload.price || 0);
  const rawOverbookingValue = payload.overbooking_allowance != null ? payload.overbooking_allowance : payload.overbookingAllowance;
  const overbookingRaw = String(rawOverbookingValue == null ? '' : rawOverbookingValue).trim();
  const overbookingAllowance = overbookingRaw === '' ? null : Number(overbookingRaw);
  if (!date) throw new Error('Date is required.');
  if (!roomTypeId) throw new Error('Room type is required.');
  if (!Number.isFinite(Number(overridePrice)) || Number(overridePrice) <= 0) {
    throw new Error('New price must be greater than zero.');
  }
  if (overbookingRaw !== '' && (!Number.isFinite(Number(overbookingAllowance)) || Number(overbookingAllowance) < 0 || !Number.isInteger(Number(overbookingAllowance)))) {
    throw new Error('Overbooking allowance must be a whole-room number of zero or higher.');
  }
  return {
    date: date,
    roomTypeId: roomTypeId,
    roomTypeName: getRoomTypeNameById_(roomTypeId),
    overridePrice: Number(overridePrice),
    overbookingAllowance: overbookingAllowance,
    note: String(payload.note || '').trim()
  };
}

function buildRateCalendarData_(startDate, days, options) {
  const readDataStartedAt = Date.now();
  const config = options || {};
  const normalizedStart = normalizeDateInput_(startDate) || stripTime_(new Date());
  const requestedDays = Number(days || 14);
  const normalizedDays = config.mode === 'month'
    ? Math.max(28, Math.min(31, isFinite(requestedDays) ? requestedDays : getDaysInMonth_(normalizedStart)))
    : (requestedDays === 30 ? 30 : 14);
  const dates = [];
  for (let index = 0; index < normalizedDays; index += 1) {
    dates.push(addDays_(normalizedStart, index));
  }
  const endExclusive = addDays_(normalizedStart, normalizedDays);
  const readBaseStartedAt = Date.now();
  const currency = getDefaultBookingCurrency_();
  const roomTypes = getActiveRoomTypeCatalog_().map(function(roomType) {
    return {
      roomTypeId: roomType.roomTypeId,
      roomTypeName: roomType.roomTypeName,
      inventoryTotal: Number(roomType.inventoryTotal || 0),
      maxGuests: Number(roomType.maxGuests || 0)
    };
  });
  const baseRateRows = getBaseRateRows_({ activeOnly: true });
  const legacyRateRows = getSheetObjects_(SHEET_NAMES.RATES).filter(function(row) {
    return String(row.status || '').trim().toLowerCase() === 'open';
  });
  logTiming_('rateCalendar:readBase', readBaseStartedAt);

  const controlsStartedAt = Date.now();
  const commercialControls = getCommercialControlRows_({ activeOnly: true });
  logTiming_('rateCalendar:readCommercialControls', controlsStartedAt);

  const inventoryStartedAt = Date.now();
  const bookingCountMap = buildConfirmedBookingRoomCountMap_(getSheetObjects_(SHEET_NAMES.BOOKINGS), {
    startDate: normalizedStart,
    endDateExclusive: endExclusive
  }).byDateRoom;
  const blockedCountMap = buildBlockedDateQtyMap_(getSheetObjects_(SHEET_NAMES.BLOCKED_DATES), {
    startDate: normalizedStart,
    endDateExclusive: endExclusive
  }).byDateRoom;
  logTiming_('rateCalendar:readInventory', inventoryStartedAt);
  logTiming_('rateCalendar:readData', readDataStartedAt);

  const cellsStartedAt = Date.now();
  const cells = {};
  roomTypes.forEach(function(roomType) {
    const roomLegacyRateRows = legacyRateRows.filter(function(row) {
      return String(row.room_type_id || '').trim() === roomType.roomTypeId;
    });
    dates.forEach(function(date) {
      const guests = Number(INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE[roomType.roomTypeId] || roomType.maxGuests || 2);
      const resolved = resolveLiveCommercialDate_(date, roomType.roomTypeId, guests, {
        commercialControls: commercialControls,
        baseRateRows: baseRateRows,
        legacyRateRows: roomLegacyRateRows,
        physicalInventory: Number(roomType.inventoryTotal || 0),
        bookingCountMap: bookingCountMap,
        blockedCountMap: blockedCountMap
      });
      const priceControl = getCommercialPriceControlForDate_(roomType.roomTypeId, date, commercialControls);
      const overbookingControl = getCommercialOverbookingControlForDate_(roomType.roomTypeId, date, commercialControls);
      const key = formatDateKey_(date) + '|' + roomType.roomTypeId;
      cells[key] = buildRateCalendarCell_(date, roomType, resolved, priceControl, overbookingControl, currency);
    });
  });
  logTiming_('rateCalendar:buildCells', cellsStartedAt);

  return {
    ok: true,
    mode: config.mode === 'month' ? 'month' : 'days',
    startDate: formatDateKey_(normalizedStart),
    endDate: formatDateKey_(addDays_(normalizedStart, normalizedDays - 1)),
    days: normalizedDays,
    month: config.month || '',
    minMonth: config.minMonth || '',
    maxMonth: config.maxMonth || '',
    dates: dates.map(function(date) { return formatDateKey_(date); }),
    roomTypes: roomTypes,
    currency: currency,
    cells: cells
  };
}

function buildRateCalendarCell_(date, roomType, resolved, priceControl, overbookingControl, currency) {
  const hasOverride = !!priceControl;
  const publicReferenceRate = resolved && resolved.publicReferenceRate == null ? null : Number(resolved && resolved.publicReferenceRate || 0);
  const savingsOffer = {
    publicReferenceRate: publicReferenceRate,
    finalNightlyRate: resolved && resolved.finalNightlyRate == null ? null : Number(resolved.finalNightlyRate || 0),
    directDiscountType: String(resolved && resolved.directDiscountType || 'None').trim(),
    directDiscountValue: resolved && resolved.directDiscountValue === '' ? '' : resolved && resolved.directDiscountValue,
    savingsAmount: Number(resolved && resolved.savingsAmount || 0),
    savingsPercentage: Number(resolved && resolved.savingsPercentage || 0),
    currency: currency
  };
  return {
    date: formatDateKey_(date),
    roomTypeId: roomType.roomTypeId,
    price: resolved && resolved.finalNightlyRate == null ? null : Number(resolved.finalNightlyRate || 0),
    basePrice: resolved && resolved.baseRate == null ? null : Number(resolved.baseRate || 0),
    comparisonPrice: publicReferenceRate,
    savingsLabel: buildDirectBookingSavingsLabel_(savingsOffer, currency),
    available: Number(resolved && resolved.availableRooms || 0),
    totalRooms: Number(roomType.inventoryTotal || resolved && resolved.physicalInventory || 0),
    soldRooms: Number(resolved && resolved.soldRooms || 0),
    blockedRooms: Number(resolved && resolved.blockedRooms || 0),
    overbookingAllowance: Number(overbookingControl && overbookingControl.overbookingAllowance || resolved && resolved.overbookingAllowanceApplied || 0),
    hasOverride: hasOverride,
    overrideType: hasOverride ? (normalizeCommercialRuleType_(priceControl.ruleType) === 'special' ? 'Special' : 'Seasonal') : '',
    source: hasOverride ? formatCommercialRuleTypeLabel_(priceControl.ruleType) : 'Base rate',
    pricingSource: String(resolved && resolved.pricingSource || 'base_rate').trim(),
    pricingReferenceId: String(resolved && resolved.pricingReferenceId || '').trim()
  };
}

function buildRateCalendarCellDataForDateRoom_(date, roomTypeId) {
  const targetDate = normalizeDateInput_(date);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const roomType = getActiveRoomTypeCatalog_().filter(function(row) {
    return row.roomTypeId === targetRoomTypeId;
  })[0];
  if (!roomType) return null;
  const nextDate = addDays_(targetDate, 1);
  const currency = getDefaultBookingCurrency_();
  const commercialControls = getCommercialControlRows_({ activeOnly: true, roomTypeId: targetRoomTypeId });
  const bookingCountMap = buildConfirmedBookingRoomCountMap_(getSheetObjects_(SHEET_NAMES.BOOKINGS), {
    startDate: targetDate,
    endDateExclusive: nextDate
  }).byDateRoom;
  const blockedCountMap = buildBlockedDateQtyMap_(getSheetObjects_(SHEET_NAMES.BLOCKED_DATES), {
    startDate: targetDate,
    endDateExclusive: nextDate
  }).byDateRoom;
  const guests = Number(INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE[targetRoomTypeId] || roomType.maxGuests || 2);
  const resolved = resolveLiveCommercialDate_(targetDate, targetRoomTypeId, guests, {
    commercialControls: commercialControls,
    baseRateRows: getBaseRateRows_({ activeOnly: true, roomTypeId: targetRoomTypeId }),
    legacyRateRows: getOpenRateRowsForType_(targetRoomTypeId),
    physicalInventory: Number(roomType.inventoryTotal || 0),
    bookingCountMap: bookingCountMap,
    blockedCountMap: blockedCountMap
  });
  return buildRateCalendarCell_(
    targetDate,
    roomType,
    resolved,
    getCommercialPriceControlForDate_(targetRoomTypeId, targetDate, commercialControls),
    getCommercialOverbookingControlForDate_(targetRoomTypeId, targetDate, commercialControls),
    currency
  );
}

function buildRateBoardCommercialSignals_(selectedDate) {
  const targetDate = normalizeFrontDeskDate_(selectedDate);
  const roomTypes = getRateBoardRoomTypes_();
  const rows = roomTypes.map(function(roomType) {
    const recommendation = buildPricingRecommendationData_(targetDate, roomType.roomTypeId);
    const pickup = recommendation && recommendation.pickupMetrics ? recommendation.pickupMetrics : {};
    const competitor = recommendation && recommendation.competitorMetrics ? recommendation.competitorMetrics : {};
    const pressure = buildRateBoardPressureState_(recommendation);
    return {
      date: formatDateKey_(targetDate),
      roomTypeId: roomType.roomTypeId,
      signalStatus: 'loaded',
      demandBand: String(recommendation && recommendation.demandBand || 'Low'),
      demandScore: Number(recommendation && recommendation.demandScore || 0),
      pickupBand: String(pickup && pickup.pickupBand || ''),
      pickupSignalRooms: pickup && pickup.pickupSignalRooms == null ? null : Number(pickup.pickupSignalRooms || 0),
      competitorAveragePrice: competitor && competitor.competitorAveragePrice == null ? null : Number(competitor.competitorAveragePrice),
      totalCompetitorsChecked: Number(competitor && competitor.totalCompetitorsChecked || 0),
      recommendationAction: String(recommendation && recommendation.recommendationAction || 'Review manually'),
      recommendationConfidence: String(recommendation && recommendation.recommendationConfidence || 'Low'),
      pressureTone: pressure.tone,
      pressureLabel: pressure.label,
      pressureMeta: pressure.meta,
      pressureRank: pressure.rank
    };
  });

  return {
    ok: true,
    selectedDate: formatDateKey_(targetDate),
    signalStatus: 'loaded',
    loadedAtLabel: Utilities.formatDate(new Date(), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm'),
    rows: rows,
    summary: {
      highPressureCount: rows.filter(function(row) { return row.pressureTone === 'red'; }).length,
      watchPressureCount: rows.filter(function(row) { return row.pressureTone === 'amber'; }).length
    }
  };
}

function buildBaseRatesData_() {
  const roomTypes = getActiveRoomTypeCatalog_();
  const rows = getBaseRateRows_({ activeOnly: false });
  return {
    ok: true,
    rates: roomTypes.map(function(roomType) {
      const match = rows.find(function(row) {
        return row.roomTypeId === roomType.roomTypeId;
      });
      const fallbackRate = getLegacyFallbackBaseRateForRoomType_(roomType.roomTypeId);
      return {
        roomTypeId: roomType.roomTypeId,
        roomTypeName: roomType.roomTypeName,
        baseRate: match ? match.baseRate : fallbackRate,
        extraGuestFee: match
          ? (match.hasExtraGuestFee ? match.extraGuestFee : '')
          : (function() {
              const legacyFallback = getLegacyRateFallbackInfo_(roomType.roomTypeId, new Date());
              return legacyFallback.found ? roundCurrency_(legacyFallback.extraGuestFee || 0) : '';
            })(),
        publicReferencePrice: match ? match.publicReferencePrice : '',
        directDiscountType: match ? match.directDiscountType : 'None',
        directDiscountValue: match ? match.directDiscountValue : '',
        active: match ? match.active : true,
        updatedAtLabel: match && match.updatedAt ? Utilities.formatDate(new Date(match.updatedAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm') : ''
      };
    })
  };
}

function getHomepageBaseRateSummary_() {
  const currency = getDefaultBookingCurrency_();
  const baseRates = buildBaseRatesData_();
  return {
    currency: currency,
    rates: (baseRates.rates || []).map(function(rate) {
      return buildHomepageBaseRateSummaryRow_(rate, currency);
    })
  };
}

function buildHomepageBaseRateSummaryRow_(rate, currency) {
  const row = rate || {};
  const roomTypeId = String(row.roomTypeId || '').trim();
  const baseRate = Number(row.baseRate || 0) > 0 ? roundCurrency_(row.baseRate) : '';
  const discount = calculateDirectBookingDiscountedBaseRate_({
    publicReferencePrice: row.publicReferencePrice,
    directDiscountType: row.directDiscountType,
    directDiscountValue: row.directDiscountValue
  }, baseRate);
  const finalDirectPrice = discount
    ? discount.directBaseRate
    : (baseRate === '' ? '' : baseRate);

  return {
    room_type_id: roomTypeId,
    room_type_name: String(row.roomTypeName || roomTypeId).trim(),
    homepage_room_name: HOMEPAGE_ROOM_DISPLAY_NAMES_BY_TYPE[roomTypeId] || '',
    base_rate: baseRate,
    public_reference_price: row.publicReferencePrice === '' || row.publicReferencePrice == null
      ? ''
      : roundCurrency_(row.publicReferencePrice),
    direct_discount_type: row.directDiscountType || 'None',
    direct_discount_value: row.directDiscountValue === '' || row.directDiscountValue == null
      ? ''
      : row.directDiscountValue,
    final_direct_price: finalDirectPrice,
    currency: currency,
    active: row.active ? 'Yes' : 'No'
  };
}

function getRateBoardRoomTypes_() {
  const orderMap = RATE_BOARD_SORT_ORDER.reduce(function(map, roomTypeId, index) {
    map[roomTypeId] = index;
    return map;
  }, {});
  return getActiveRoomTypeCatalog_().slice().sort(function(a, b) {
    const aOrder = Object.prototype.hasOwnProperty.call(orderMap, a.roomTypeId) ? orderMap[a.roomTypeId] : 999;
    const bOrder = Object.prototype.hasOwnProperty.call(orderMap, b.roomTypeId) ? orderMap[b.roomTypeId] : 999;
    if (aOrder !== bOrder) return aOrder - bOrder;
    return String(a.roomTypeName || '').localeCompare(String(b.roomTypeName || ''));
  });
}

function getRateLadderConfig_() {
  const defaults = buildDefaultRateLadderConfig_();
  const activeRoomTypes = getRateBoardRoomTypes_();
  const activeIds = activeRoomTypes.map(function(roomType) { return roomType.roomTypeId; });
  let stored = null;

  try {
    const raw = PropertiesService.getScriptProperties().getProperty(RATE_LADDER_PROPERTY_KEY);
    stored = raw ? JSON.parse(raw) : null;
  } catch (error) {
    stored = null;
  }

  const leadRoomTypeId = stored && activeIds.indexOf(stored.leadRoomTypeId) !== -1
    ? stored.leadRoomTypeId
    : defaults.leadRoomTypeId;
  const differentials = {};

  activeIds.forEach(function(roomTypeId) {
    const storedValue = stored && stored.differentials ? stored.differentials[roomTypeId] : null;
    const fallbackValue = defaults.differentials[roomTypeId] != null ? defaults.differentials[roomTypeId] : 0;
    differentials[roomTypeId] = Number.isFinite(Number(storedValue)) ? roundCurrency_(storedValue) : roundCurrency_(fallbackValue);
  });

  return {
    leadRoomTypeId: leadRoomTypeId,
    leadRoomTypeName: getRoomTypeNameById_(leadRoomTypeId),
    differentials: differentials,
    updatedAtLabel: stored && stored.updatedAt
      ? Utilities.formatDate(new Date(stored.updatedAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm')
      : '',
    lastAppliedAtLabel: stored && stored.lastAppliedAt
      ? Utilities.formatDate(new Date(stored.lastAppliedAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm')
      : '',
    lastReason: String(stored && stored.lastReason || '').trim()
  };
}

function buildDefaultRateLadderConfig_() {
  const roomTypes = getRateBoardRoomTypes_();
  const rates = buildBaseRatesData_().rates || [];
  const rateMap = {};
  rates.forEach(function(rate) {
    rateMap[rate.roomTypeId] = Number(rate.baseRate || 0);
  });

  const defaultLeadRoomTypeId = roomTypes.some(function(roomType) {
    return roomType.roomTypeId === RATE_LADDER_LEAD_ROOM_TYPE_ID;
  }) ? RATE_LADDER_LEAD_ROOM_TYPE_ID : (roomTypes[0] ? roomTypes[0].roomTypeId : '');
  const leadBaseRate = Number(rateMap[defaultLeadRoomTypeId] || 0);
  const differentials = {};

  roomTypes.forEach(function(roomType) {
    differentials[roomType.roomTypeId] = roundCurrency_(Number(rateMap[roomType.roomTypeId] || 0) - leadBaseRate);
  });

  return {
    leadRoomTypeId: defaultLeadRoomTypeId,
    differentials: differentials
  };
}

function saveRateLadderConfig_(config) {
  if (!config || !config.leadRoomTypeId) return;
  PropertiesService.getScriptProperties().setProperty(RATE_LADDER_PROPERTY_KEY, JSON.stringify({
    leadRoomTypeId: config.leadRoomTypeId,
    differentials: config.differentials || {},
    updatedAt: new Date().toISOString(),
    lastReason: String(config.lastReason || '').trim(),
    lastAppliedAt: String(config.lastAppliedAt || '').trim()
  }));
}

function validateRateLadderInput_(input) {
  const payload = input || {};
  const leadBaseRate = roundCurrency_(payload.lead_base_rate || payload.leadBaseRate || payload.base_rate || payload.baseRate || 0);
  const changeReason = String(payload.change_reason || payload.changeReason || payload.operator_note || payload.operatorNote || '').trim();
  if (!Number.isFinite(Number(leadBaseRate)) || Number(leadBaseRate) <= 0) {
    throw new Error('Lead base rate must be greater than zero.');
  }
  if (!changeReason) {
    throw new Error('Add an operator reason before applying the rate ladder.');
  }
  return {
    leadBaseRate: Number(leadBaseRate),
    changeReason: changeReason,
    selectedDate: normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date())
  };
}

function buildRateBoardPressureState_(recommendation) {
  const occupancyPct = Number(recommendation && recommendation.occupancyPct || 0);
  const exact = recommendation && recommendation.exactDateMetrics ? recommendation.exactDateMetrics : {};
  const availableRooms = Number(exact.availableRooms || 0);
  const pressureBand = recommendation && recommendation.next7DayPressure ? String(recommendation.next7DayPressure.band || '') : '';
  const soldRooms = Number(exact.soldRooms || 0);
  const sellableInventory = Number(exact.sellableInventory || 0);
  let tone = 'green';
  let label = 'Green';

  if (availableRooms <= 1 || occupancyPct >= 0.85 || pressureBand === 'Hot') {
    tone = 'red';
    label = 'Red';
  } else if (availableRooms <= 2 || occupancyPct >= 0.6 || pressureBand === 'Strong' || pressureBand === 'Normal') {
    tone = 'amber';
    label = 'Amber';
  }

  return {
    tone: tone,
    label: label,
    rank: tone === 'red' ? 3 : (tone === 'amber' ? 2 : 1),
    meta: soldRooms + ' sold of ' + sellableInventory + ' | ' + formatPercentValue_(occupancyPct)
  };
}

function buildRateBoardOverrideState_(roomTypeId, leadRoomTypeId, followsLadder, priceControl) {
  const hasBaseBreakaway = roomTypeId !== leadRoomTypeId && !followsLadder;
  const states = [];
  if (priceControl && priceControl.hasOverridePrice) {
    states.push({
      tone: 'red',
      label: 'Date override',
      meta: formatCommercialRuleTypeLabel_(priceControl.ruleType) + ' is active on this date.'
    });
  }
  if (hasBaseBreakaway) {
    states.push({
      tone: 'amber',
      label: 'Broken from ladder',
      meta: 'Base rate has diverged from the stored ladder differential.'
    });
  }
  if (!states.length) {
    states.push({
      tone: 'green',
      label: roomTypeId === leadRoomTypeId ? 'Lead rate' : 'Following ladder',
      meta: roomTypeId === leadRoomTypeId ? 'This room type drives the ladder.' : 'Base rate follows the ladder differential.'
    });
  }
  return {
    tone: states.some(function(state) { return state.tone === 'red'; }) ? 'red' : (states.some(function(state) { return state.tone === 'amber'; }) ? 'amber' : 'green'),
    label: states[0].label,
    meta: states.map(function(state) { return state.meta; }).join(' '),
    hasDateOverride: !!(priceControl && priceControl.hasOverridePrice),
    hasLadderBreakaway: hasBaseBreakaway,
    states: states
  };
}

function getLatestOtaWorkflowStateForDate_(rows, targetDate, roomTypeId) {
  const targetKey = formatDateKey_(targetDate);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const match = (rows || [])
    .filter(function(row) {
      return row.date && formatDateKey_(row.date) === targetKey && row.roomTypeId === targetRoomTypeId;
    })
    .sort(compareOtaWorkflowRows_)[0];

  if (!match) {
    return {
      status: 'Not queued',
      meta: 'No OTA follow-up row logged yet.',
      updatedAtLabel: ''
    };
  }

  const bookingStatus = String(match.bookingComStatus || '').trim() || 'Pending';
  const airbnbStatus = String(match.airbnbStatus || '').trim() || 'Pending';
  let status = 'Pending';
  if (bookingStatus === 'Updated' && airbnbStatus === 'Updated') {
    status = 'Updated';
  } else if (bookingStatus === 'Not Relevant' && airbnbStatus === 'Not Relevant') {
    status = 'Not relevant';
  } else if (bookingStatus === 'Updated' || airbnbStatus === 'Updated') {
    status = 'Mixed';
  }

  return {
    status: status,
    meta: 'B.com ' + bookingStatus + ' | Airbnb ' + airbnbStatus,
    updatedAtLabel: match.updatedAt ? Utilities.formatDate(new Date(match.updatedAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm') : ''
  };
}

function buildPmsOperatorStatus_(options) {
  const config = options || {};
  const contact = config.contactDetails || getWebsiteContactDetails_();
  let remainingDailyQuota = null;
  let mailQuotaState = config.skipMailQuota ? 'deferred' : 'unknown';

  if (!config.skipMailQuota) {
    try {
      remainingDailyQuota = MailApp.getRemainingDailyQuota();
      mailQuotaState = 'ok';
    } catch (error) {
      mailQuotaState = 'unavailable';
    }
  }

  return {
    notificationEmail: String(contact.notificationEmail || '').trim(),
    notificationEmailConfigured: !!String(contact.notificationEmail || '').trim(),
    latestBookingCreatedAt: config.latestBookingCreatedAt || getLatestBookingAlertCursor_(config.bookings),
    remainingDailyQuota: remainingDailyQuota,
    mailQuotaState: mailQuotaState
  };
}

function getLatestBookingAlertCursor_(bookingRows) {
  let latest = null;
  const rows = bookingRows || getBookingRowsForAlertCursor_();
  rows.forEach(function(row) {
    const createdAt = normalizeDateTimeInput_(row.created_at);
    if (!createdAt) return;
    if (!latest || createdAt.getTime() > latest.getTime()) {
      latest = createdAt;
    }
  });
  return latest ? latest.toISOString() : '';
}

function extractBookingAlertRecoveryState_(row) {
  const notes = String(row && row.internal_notes || '').trim();
  if (!notes) {
    return {
      recoveryNeeded: false,
      recoveryNote: ''
    };
  }
  const match = notes.match(/Recovery needed after booking creation:\s*([^\r\n]+)/i);
  if (!match) {
    return {
      recoveryNeeded: false,
      recoveryNote: ''
    };
  }
  return {
    recoveryNeeded: true,
    recoveryNote: String(match[1] || '').trim()
  };
}

function serializeRecentBookingAlertRow_(row) {
  const createdAt = normalizeDateTimeInput_(row.created_at);
  const recoveryState = extractBookingAlertRecoveryState_(row);
  return {
    bookingId: String(row.booking_id || '').trim(),
    guestName: String(row.guest_name || '').trim(),
    source: String(row.source || '').trim(),
    roomTypeName: String(row.room_type_name || '').trim(),
    checkIn: normalizeDateInput_(row.check_in) ? formatDateKey_(row.check_in) : '',
    checkOut: normalizeDateInput_(row.check_out) ? formatDateKey_(row.check_out) : '',
    balanceDue: Number(row.balance_due || 0),
    status: String(row.status || '').trim(),
    createdAt: createdAt,
    createdAtIso: createdAt ? createdAt.toISOString() : '',
    createdAtLabel: createdAt ? Utilities.formatDate(createdAt, getScriptTimeZone_(), 'yyyy-MM-dd HH:mm') : '',
    recoveryNeeded: recoveryState.recoveryNeeded,
    recoveryNote: recoveryState.recoveryNote
  };
}


function normalizeEventFlagType_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const match = EVENT_FLAG_TYPE_VALUES.find(function(option) {
    return String(option || '').trim().toLowerCase() === raw.toLowerCase();
  });
  return match || raw;
}

function normalizeEventFlagImpactLevel_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const match = EVENT_FLAG_IMPACT_VALUES.find(function(option) {
    return String(option || '').trim().toLowerCase() === raw.toLowerCase();
  });
  return match || raw;
}

function getEventFlagImpactRank_(impactLevel) {
  const normalized = String(impactLevel || '').trim().toLowerCase();
  if (normalized === 'very high') return 4;
  if (normalized === 'high') return 3;
  if (normalized === 'medium') return 2;
  if (normalized === 'low') return 1;
  return 0;
}

function getEventFlagImpactDemandComponent_(impactLevel) {
  const rank = getEventFlagImpactRank_(impactLevel);
  if (rank >= 4) return 0.85;
  if (rank === 3) return 0.65;
  if (rank === 2) return 0.4;
  if (rank === 1) return 0.2;
  return null;
}

function datesOverlapInclusive_(startA, endA, startB, endB) {
  if (!startA || !endA || !startB || !endB) return false;
  return startA.getTime() <= endB.getTime() && endA.getTime() >= startB.getTime();
}

function isDateWithinInclusiveRange_(date, startDate, endDate) {
  if (!date || !startDate || !endDate) return false;
  const target = stripTime_(date).getTime();
  return target >= stripTime_(startDate).getTime() && target <= stripTime_(endDate).getTime();
}

function normalizeEventFlagRow_(row, rowNumber) {
  const startDate = normalizeDateInput_(row.start_date || row.date);
  const endDate = normalizeDateInput_(row.end_date || row.start_date || row.date || row.startDate) || startDate;
  const eventType = normalizeEventFlagType_(row.event_type);
  const impactLevel = normalizeEventFlagImpactLevel_(row.impact_level);
  const active = toBoolean_(row.active);

  return {
    rowNumber: Number(rowNumber || 0),
    eventId: String(row.event_id || '').trim(),
    eventName: String(row.event_name || '').trim(),
    startDate: startDate,
    endDate: endDate,
    eventType: eventType,
    impactLevel: impactLevel,
    impactRank: getEventFlagImpactRank_(impactLevel),
    note: String(row.note || '').trim(),
    active: active,
    updatedAt: row.updated_at || '',
    updatedAtMs: toDateTimeMs_(row.updated_at)
  };
}

function serializeEventFlagRow_(row) {
  const startDateKey = row.startDate ? formatDateKey_(row.startDate) : '';
  const endDateKey = row.endDate ? formatDateKey_(row.endDate) : '';
  return {
    eventId: row.eventId,
    eventName: row.eventName,
    startDate: startDateKey,
    endDate: endDateKey,
    dateRangeLabel: startDateKey && endDateKey && startDateKey !== endDateKey ? startDateKey + ' to ' + endDateKey : startDateKey,
    eventType: row.eventType,
    impactLevel: row.impactLevel,
    note: row.note,
    active: row.active,
    activeLabel: row.active ? 'Active' : 'Inactive',
    updatedAtLabel: row.updatedAtMs
      ? Utilities.formatDate(new Date(row.updatedAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm')
      : ''
  };
}

function compareEventFlagRows_(a, b) {
  if (!!a.active !== !!b.active) return a.active ? -1 : 1;
  const startDelta = Number(a && a.startDate ? a.startDate.getTime() : 0) - Number(b && b.startDate ? b.startDate.getTime() : 0);
  if (startDelta !== 0) return startDelta;
  const impactDelta = Number(b && b.impactRank || 0) - Number(a && a.impactRank || 0);
  if (impactDelta !== 0) return impactDelta;
  const updatedDelta = Number(b && b.updatedAtMs || 0) - Number(a && a.updatedAtMs || 0);
  if (updatedDelta !== 0) return updatedDelta;
  return String(a && a.eventName || '').localeCompare(String(b && b.eventName || ''));
}

function validateEventFlagFilters_(input) {
  const payload = input || {};
  const fallbackDate = normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date());
  const selectedDate = normalizeDateInput_(payload.selected_date || payload.selectedDate || payload.date || fallbackDate) || fallbackDate;
  const startDate = normalizeDateInput_(payload.start_date || payload.startDate || selectedDate) || selectedDate;
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate || addDays_(startDate, 90)) || addDays_(startDate, 90);
  if (endDate.getTime() < startDate.getTime()) {
    throw new Error('End date must be on or after start date.');
  }

  return {
    selectedDate: selectedDate,
    startDate: startDate,
    endDate: endDate,
    includeInactive: toBoolean_(payload.include_inactive || payload.includeInactive)
  };
}

function serializeEventFlagFilters_(filters) {
  return {
    selectedDate: formatDateKey_(filters.selectedDate),
    startDate: formatDateKey_(filters.startDate),
    endDate: formatDateKey_(filters.endDate),
    includeInactive: !!filters.includeInactive
  };
}

function validateEventFlagEntryInput_(input) {
  const payload = input || {};
  const eventName = String(payload.event_name || payload.eventName || '').trim();
  if (!eventName) throw new Error('Event name is required.');

  const startDate = normalizeDateInput_(payload.start_date || payload.startDate || payload.date);
  if (!startDate) throw new Error('Start date is required.');
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate || payload.start_date || payload.startDate || payload.date) || startDate;
  if (endDate.getTime() < startDate.getTime()) {
    throw new Error('End date must be on or after start date.');
  }

  const eventType = normalizeEventFlagType_(payload.event_type || payload.eventType);
  if (!eventType || EVENT_FLAG_TYPE_VALUES.indexOf(eventType) === -1) {
    throw new Error('Event type is invalid.');
  }

  const impactLevel = normalizeEventFlagImpactLevel_(payload.impact_level || payload.impactLevel);
  if (!impactLevel || EVENT_FLAG_IMPACT_VALUES.indexOf(impactLevel) === -1) {
    throw new Error('Impact level is invalid.');
  }

  const eventId = String(payload.event_id || payload.eventId || '').trim();
  let rowNumber = 0;
  if (eventId) {
    const sheet = ensureSheetWithHeaders_(getSpreadsheet_(), SHEET_NAMES.EVENT_FLAGS, EVENT_FLAGS_HEADERS);
    rowNumber = findRowNumberByHeaderValue_(sheet, 'event_id', eventId);
    if (!rowNumber) {
      throw new Error('Event flag not found: ' + eventId);
    }
  }

  return {
    eventId: eventId,
    rowNumber: rowNumber,
    eventName: eventName,
    startDate: startDate,
    endDate: endDate,
    eventType: eventType,
    impactLevel: impactLevel,
    note: String(payload.note || '').trim(),
    active: !Object.prototype.hasOwnProperty.call(payload, 'active') && !Object.prototype.hasOwnProperty.call(payload, 'is_active')
      ? true
      : toBoolean_(payload.active || payload.is_active || payload.isActive)
  };
}

function buildEventFlagSelectedDateSummary_(selectedDate, rows) {
  const normalizedDate = normalizeDateInput_(selectedDate) || normalizeFrontDeskDate_(new Date());
  const matchingRows = (rows || []).filter(function(row) {
    return row && row.active && isDateWithinInclusiveRange_(normalizedDate, row.startDate, row.endDate);
  }).sort(function(a, b) {
    const impactDelta = Number(b && b.impactRank || 0) - Number(a && a.impactRank || 0);
    if (impactDelta !== 0) return impactDelta;
    return compareEventFlagRows_(a, b);
  });
  const highestImpactLevel = matchingRows.length ? String(matchingRows[0].impactLevel || '') : '';

  return {
    selectedDate: formatDateKey_(normalizedDate),
    hasActiveFlag: matchingRows.length > 0,
    highestImpactLevel: highestImpactLevel,
    highestImpactRank: getEventFlagImpactRank_(highestImpactLevel),
    highestImpactScore: getEventFlagImpactDemandComponent_(highestImpactLevel),
    flagCount: matchingRows.length,
    eventNames: matchingRows.map(function(row) { return row.eventName; }),
    flags: matchingRows.map(serializeEventFlagRow_)
  };
}

function getCompetitorTrackerRoomTypeLabel_(roomTypeId, roomTypeName) {
  const fallbackName = roomTypeId ? getRoomTypeNameById_(roomTypeId) : '';
  const name = String(roomTypeName || fallbackName || '').trim();
  if (/shared bathroom/i.test(name)) return 'Classic';
  if (/with balcony/i.test(name)) return 'Peak Mountain View with Balcony';
  if (/cottage/i.test(name)) return 'Peak Cottage';
  if (/peak mountain view/i.test(name)) return 'Peak Mountain View';
  return name || String(roomTypeId || '').trim();
}

function normalizeCompetitorTrackerAvailability_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  return normalized === 'sold out' || normalized === 'sold_out' || normalized === 'soldout'
    ? 'Sold Out'
    : 'Available';
}

function normalizeCompetitorTrackerBreakfast_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  if (normalized === 'yes' || normalized === 'y' || normalized === 'included') return 'Yes';
  if (normalized === 'no' || normalized === 'n' || normalized === 'not included') return 'No';
  return 'Unknown';
}

function normalizeCompetitorTrackerSource_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const normalized = raw.toLowerCase();
  if (normalized === 'booking.com' || normalized === 'booking') return 'Booking.com';
  if (normalized === 'airbnb') return 'Airbnb';
  if (normalized === 'direct website' || normalized === 'direct') return 'Direct website';
  if (normalized === 'other') return 'Other';
  return raw;
}

function normalizeCompetitorTrackerRow_(row, rowNumber) {
  const date = normalizeDateInput_(row.date);
  const roomTypeId = String(row.matched_roza_room_type_id || row.room_type_id || '').trim();
  const competitorPriceRaw = String(row.competitor_price == null ? '' : row.competitor_price).trim();
  const competitorPrice = competitorPriceRaw === '' ? null : roundCurrency_(row.competitor_price);
  return {
    rowNumber: Number(rowNumber || 0),
    entryId: String(row.entry_id || '').trim(),
    createdAt: row.created_at || '',
    createdAtMs: toDateTimeMs_(row.created_at),
    date: date,
    dateKey: date ? formatDateKey_(date) : '',
    competitorName: String(row.competitor_name || '').trim(),
    matchedRozaRoomTypeId: roomTypeId,
    matchedRozaRoomType: getCompetitorTrackerRoomTypeLabel_(roomTypeId, row.matched_roza_room_type),
    competitorRoomDescription: String(row.competitor_room_description || '').trim(),
    competitorPrice: competitorPrice,
    availableOrSoldOut: normalizeCompetitorTrackerAvailability_(row.available_or_sold_out),
    breakfastIncluded: normalizeCompetitorTrackerBreakfast_(row.breakfast_included),
    sourceChecked: normalizeCompetitorTrackerSource_(row.source_checked),
    note: String(row.note || '').trim()
  };
}

function serializeCompetitorTrackerRow_(row) {
  return {
    entryId: row.entryId,
    date: row.dateKey,
    competitorName: row.competitorName,
    matchedRozaRoomTypeId: row.matchedRozaRoomTypeId,
    matchedRozaRoomType: row.matchedRozaRoomType,
    competitorRoomDescription: row.competitorRoomDescription,
    competitorPrice: row.competitorPrice,
    availableOrSoldOut: row.availableOrSoldOut,
    breakfastIncluded: row.breakfastIncluded,
    sourceChecked: row.sourceChecked,
    note: row.note,
    createdAtLabel: row.createdAtMs
      ? Utilities.formatDate(new Date(row.createdAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm')
      : ''
  };
}

function compareCompetitorTrackerRows_(a, b) {
  const dateDelta = Number(b && b.date ? b.date.getTime() : 0) - Number(a && a.date ? a.date.getTime() : 0);
  if (dateDelta !== 0) return dateDelta;
  const roomDelta = String(a && a.matchedRozaRoomType || '').localeCompare(String(b && b.matchedRozaRoomType || ''));
  if (roomDelta !== 0) return roomDelta;
  const competitorDelta = String(a && a.competitorName || '').localeCompare(String(b && b.competitorName || ''));
  if (competitorDelta !== 0) return competitorDelta;
  return Number(b && b.createdAtMs || 0) - Number(a && a.createdAtMs || 0);
}

function validateCompetitorTrackerFilters_(input) {
  const payload = input || {};
  const fallbackDate = normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date());
  const startDate = normalizeDateInput_(payload.start_date || payload.startDate || payload.date || fallbackDate) || fallbackDate;
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate || payload.date || startDate) || startDate;
  if (endDate.getTime() < startDate.getTime()) {
    throw new Error('End date must be on or after start date.');
  }

  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || payload.matched_roza_room_type_id || payload.matchedRozaRoomTypeId || '').trim();
  return {
    startDate: startDate,
    endDate: endDate,
    roomTypeId: roomTypeRaw ? resolveRoomTypeId_(roomTypeRaw) : ''
  };
}

function serializeCompetitorTrackerFilters_(filters) {
  return {
    startDate: formatDateKey_(filters.startDate),
    endDate: formatDateKey_(filters.endDate),
    roomTypeId: filters.roomTypeId,
    roomTypeLabel: filters.roomTypeId ? getCompetitorTrackerRoomTypeLabel_(filters.roomTypeId) : '',
    singleDate: formatDateKey_(filters.startDate) === formatDateKey_(filters.endDate)
  };
}

function validateCompetitorTrackerEntryInput_(input) {
  const payload = input || {};
  const date = normalizeDateInput_(payload.date);
  if (!date) throw new Error('Date is required.');

  const competitorName = String(payload.competitor_name || payload.competitorName || '').trim();
  if (!competitorName) throw new Error('Competitor name is required.');

  const roomTypeRaw = String(payload.matched_roza_room_type_id || payload.matchedRozaRoomTypeId || payload.matched_roza_room_type || payload.matchedRozaRoomType || payload.room_type_id || payload.roomTypeId || '').trim();
  if (!roomTypeRaw) throw new Error('Matched Roza room type is required.');
  const matchedRozaRoomTypeId = resolveRoomTypeId_(roomTypeRaw);
  const matchedRozaRoomType = getCompetitorTrackerRoomTypeLabel_(matchedRozaRoomTypeId);

  const competitorRoomDescription = String(payload.competitor_room_description || payload.competitorRoomDescription || '').trim();
  if (!competitorRoomDescription) throw new Error('Competitor room description is required.');

  const competitorPriceRaw = String(payload.competitor_price == null ? '' : payload.competitor_price).trim();
  if (competitorPriceRaw === '') throw new Error('Competitor price is required.');
  const competitorPrice = roundCurrency_(competitorPriceRaw);
  if (!isFinite(Number(competitorPrice)) || Number(competitorPrice) < 0) {
    throw new Error('Competitor price must be a valid amount.');
  }

  const availableRaw = String(payload.available_or_sold_out || payload.availableOrSoldOut || '').trim();
  if (!availableRaw) throw new Error('Availability status is required.');
  const availableOrSoldOut = normalizeCompetitorTrackerAvailability_(availableRaw);

  const breakfastRaw = String(payload.breakfast_included || payload.breakfastIncluded || '').trim();
  if (!breakfastRaw) throw new Error('Breakfast included is required.');
  const breakfastIncluded = normalizeCompetitorTrackerBreakfast_(breakfastRaw);

  const sourceRaw = String(payload.source_checked || payload.sourceChecked || '').trim();
  if (!sourceRaw) throw new Error('Source checked is required.');
  const sourceChecked = normalizeCompetitorTrackerSource_(sourceRaw);

  return {
    date: date,
    competitorName: competitorName,
    matchedRozaRoomTypeId: matchedRozaRoomTypeId,
    matchedRozaRoomType: matchedRozaRoomType,
    competitorRoomDescription: competitorRoomDescription,
    competitorPrice: competitorPrice,
    availableOrSoldOut: availableOrSoldOut,
    breakfastIncluded: breakfastIncluded,
    sourceChecked: sourceChecked,
    note: String(payload.note || '').trim()
  };
}

function normalizeRecommendationActionLogRow_(row, rowNumber) {
  const date = normalizeDateInput_(row.date);
  const roomTypeId = String(row.room_type_id || '').trim();
  const roomTypeName = String(row.room_type_name || (roomTypeId ? getRoomTypeNameById_(roomTypeId) : '') || '').trim();
  const rozaOldPriceRaw = String(row.roza_old_price == null ? '' : row.roza_old_price).trim();
  const rozaNewPriceRaw = String(row.roza_new_price == null ? '' : row.roza_new_price).trim();

  return {
    rowNumber: Number(rowNumber || 0),
    logId: String(row.log_id || '').trim(),
    date: date,
    dateKey: date ? formatDateKey_(date) : '',
    roomTypeId: roomTypeId,
    roomTypeName: roomTypeName,
    recommendationAction: String(row.recommendation_action || '').trim(),
    recommendationConfidence: String(row.recommendation_confidence || '').trim(),
    recommendationReasonSummary: String(row.recommendation_reason_summary || '').trim(),
    rozaOldPrice: rozaOldPriceRaw === '' ? null : roundCurrency_(row.roza_old_price),
    actionTaken: String(row.action_taken || '').trim(),
    rozaNewPrice: rozaNewPriceRaw === '' ? null : roundCurrency_(row.roza_new_price),
    overbookingChange: String(row.overbooking_change || '').trim(),
    otaUpdated: String(row.ota_updated || '').trim(),
    operatorNote: String(row.operator_note || '').trim(),
    createdAt: row.created_at || '',
    createdAtMs: toDateTimeMs_(row.created_at)
  };
}

function serializeRecommendationActionLogRow_(row) {
  return {
    logId: row.logId,
    date: row.dateKey,
    roomTypeId: row.roomTypeId,
    roomTypeName: row.roomTypeName,
    recommendationAction: row.recommendationAction,
    recommendationConfidence: row.recommendationConfidence,
    recommendationReasonSummary: row.recommendationReasonSummary,
    rozaOldPrice: row.rozaOldPrice,
    actionTaken: row.actionTaken,
    rozaNewPrice: row.rozaNewPrice,
    overbookingChange: row.overbookingChange,
    otaUpdated: row.otaUpdated,
    operatorNote: row.operatorNote,
    createdAtLabel: row.createdAtMs
      ? Utilities.formatDate(new Date(row.createdAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm')
      : ''
  };
}

function compareRecommendationActionLogRows_(a, b) {
  const dateDelta = Number(b && b.date ? b.date.getTime() : 0) - Number(a && a.date ? a.date.getTime() : 0);
  if (dateDelta !== 0) return dateDelta;
  const createdDelta = Number(b && b.createdAtMs || 0) - Number(a && a.createdAtMs || 0);
  if (createdDelta !== 0) return createdDelta;
  return String(a && a.roomTypeName || '').localeCompare(String(b && b.roomTypeName || ''));
}

function validateRecommendationActionLogFilters_(input) {
  const payload = input || {};
  const fallbackDate = normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date());
  const startDate = normalizeDateInput_(payload.start_date || payload.startDate || payload.date || fallbackDate) || fallbackDate;
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate || payload.date || startDate) || startDate;
  if (endDate.getTime() < startDate.getTime()) {
    throw new Error('End date must be on or after start date.');
  }

  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || '').trim();
  return {
    startDate: startDate,
    endDate: endDate,
    roomTypeId: roomTypeRaw ? resolveRoomTypeId_(roomTypeRaw) : ''
  };
}

function serializeRecommendationActionLogFilters_(filters) {
  return {
    startDate: formatDateKey_(filters.startDate),
    endDate: formatDateKey_(filters.endDate),
    roomTypeId: filters.roomTypeId,
    roomTypeLabel: filters.roomTypeId ? getCompetitorTrackerRoomTypeLabel_(filters.roomTypeId) : '',
    singleDate: formatDateKey_(filters.startDate) === formatDateKey_(filters.endDate)
  };
}

function validateRecommendationActionLogEntryInput_(input) {
  const payload = input || {};
  const date = normalizeDateInput_(payload.date || payload.selected_date || payload.selectedDate);
  if (!date) throw new Error('Date is required.');

  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || '').trim();
  if (!roomTypeRaw) throw new Error('Room type is required.');
  const roomTypeId = resolveRoomTypeId_(roomTypeRaw);

  const recommendationAction = String(payload.recommendation_action || payload.recommendationAction || '').trim();
  if (!recommendationAction) throw new Error('Recommendation action is required.');

  const actionTakenRaw = String(payload.action_taken || payload.actionTaken || '').trim();
  if (!actionTakenRaw) throw new Error('Action taken is required.');
  const actionTaken = normalizeRecommendationActionTaken_(actionTakenRaw);

  const otaUpdatedRaw = String(payload.ota_updated || payload.otaUpdated || '').trim();
  if (!otaUpdatedRaw) throw new Error('OTA updated is required.');
  const otaUpdated = normalizeRecommendationOtaUpdated_(otaUpdatedRaw);

  const overbookingChange = normalizeRecommendationOverbookingChange_(payload.overbooking_change || payload.overbookingChange || 'No change');
  const rozaOldPrice = parseRecommendationActionLogMoney_(payload.roza_old_price || payload.rozaOldPrice, 'Old price');
  const rozaNewPriceRaw = payload.roza_new_price == null ? '' : payload.roza_new_price;
  const rozaNewPrice = String(rozaNewPriceRaw).trim() === '' ? rozaOldPrice : parseRecommendationActionLogMoney_(rozaNewPriceRaw, 'New price');

  return {
    date: date,
    roomTypeId: roomTypeId,
    roomTypeName: getRoomTypeNameById_(roomTypeId),
    recommendationAction: recommendationAction,
    recommendationConfidence: String(payload.recommendation_confidence || payload.recommendationConfidence || '').trim(),
    recommendationReasonSummary: String(payload.recommendation_reason_summary || payload.recommendationReasonSummary || '').trim(),
    rozaOldPrice: rozaOldPrice,
    actionTaken: actionTaken,
    rozaNewPrice: rozaNewPrice,
    overbookingChange: overbookingChange,
    otaUpdated: otaUpdated,
    operatorNote: String(payload.operator_note || payload.operatorNote || '').trim()
  };
}

function parseRecommendationActionLogMoney_(value, label) {
  const raw = String(value == null ? '' : value).trim();
  if (raw === '') {
    throw new Error(label + ' is required.');
  }
  const parsed = roundCurrency_(raw);
  if (!isFinite(Number(parsed)) || Number(parsed) < 0) {
    throw new Error(label + ' must be a valid amount.');
  }
  return parsed;
}

function getRecommendationActionTakenOptions_() {
  const currency = getDefaultBookingCurrency_();
  return [
    'Held price',
    'Raised ' + currency + ' 5',
    'Raised ' + currency + ' 10',
    'Raised ' + currency + ' 20',
    'Dropped ' + currency + ' 5',
    'Dropped ' + currency + ' 10',
    'Raised custom',
    'Dropped custom',
    'Set exact price',
    'Changed manually',
    'No action',
    'Reviewed only'
  ];
}

function normalizeRecommendationActionTaken_(value) {
  const raw = String(value || '').trim().toLowerCase();
  const match = getRecommendationActionTakenOptions_().concat(RECOMMENDATION_ACTION_TAKEN_VALUES).find(function(option) {
    return String(option || '').trim().toLowerCase() === raw;
  });
  if (!match) throw new Error('Action taken is invalid.');
  return match;
}

function normalizeRecommendationOtaUpdated_(value) {
  const raw = String(value || '').trim().toLowerCase();
  const match = RECOMMENDATION_OTA_UPDATED_VALUES.find(function(option) {
    return String(option || '').trim().toLowerCase() === raw;
  });
  if (!match) throw new Error('OTA updated value is invalid.');
  return match;
}

function normalizeRecommendationOverbookingChange_(value) {
  const raw = String(value || 'No change').trim().toLowerCase();
  const match = RECOMMENDATION_OVERBOOKING_CHANGE_VALUES.find(function(option) {
    return String(option || '').trim().toLowerCase() === raw;
  });
  if (!match) throw new Error('Overbooking change value is invalid.');
  return match;
}

function normalizeOtaWorkflowStatus_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const normalized = raw.toLowerCase();
  if (normalized === 'updated') return 'Updated';
  if (normalized === 'pending') return 'Pending';
  if (normalized === 'not relevant' || normalized === 'not_relevant' || normalized === 'n/a') return 'Not Relevant';
  const match = OTA_WORKFLOW_STATUS_VALUES.find(function(option) {
    return String(option || '').trim().toLowerCase() === normalized;
  });
  return match || raw;
}

function normalizeOtaWorkflowRow_(row, rowNumber) {
  const date = normalizeDateInput_(row.date);
  const roomTypeId = String(row.room_type_id || '').trim();
  const roomTypeName = String(row.room_type_name || (roomTypeId ? getRoomTypeNameById_(roomTypeId) : '') || '').trim();
  const priceRaw = String(row.roza_price_after_change == null ? '' : row.roza_price_after_change).trim();

  return {
    rowNumber: Number(rowNumber || 0),
    otaUpdateId: String(row.ota_update_id || '').trim(),
    date: date,
    dateKey: date ? formatDateKey_(date) : '',
    roomTypeId: roomTypeId,
    roomTypeName: roomTypeName,
    bookingComStatus: normalizeOtaWorkflowStatus_(row.booking_com_status),
    airbnbStatus: normalizeOtaWorkflowStatus_(row.airbnb_status),
    linkedRecommendationAction: String(row.linked_recommendation_action || '').trim(),
    linkedActionTaken: String(row.linked_action_taken || '').trim(),
    rozaPriceAfterChange: priceRaw === '' ? null : roundCurrency_(row.roza_price_after_change),
    note: String(row.note || '').trim(),
    updatedAt: row.updated_at || '',
    updatedAtMs: toDateTimeMs_(row.updated_at)
  };
}

function serializeOtaWorkflowRow_(row) {
  return {
    otaUpdateId: row.otaUpdateId,
    date: row.dateKey,
    roomTypeId: row.roomTypeId,
    roomTypeName: row.roomTypeName,
    bookingComStatus: row.bookingComStatus,
    airbnbStatus: row.airbnbStatus,
    linkedRecommendationAction: row.linkedRecommendationAction,
    linkedActionTaken: row.linkedActionTaken,
    rozaPriceAfterChange: row.rozaPriceAfterChange,
    note: row.note,
    updatedAtLabel: row.updatedAtMs
      ? Utilities.formatDate(new Date(row.updatedAt), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm')
      : ''
  };
}

function compareOtaWorkflowRows_(a, b) {
  const dateDelta = Number(b && b.date ? b.date.getTime() : 0) - Number(a && a.date ? a.date.getTime() : 0);
  if (dateDelta !== 0) return dateDelta;
  const updatedDelta = Number(b && b.updatedAtMs || 0) - Number(a && a.updatedAtMs || 0);
  if (updatedDelta !== 0) return updatedDelta;
  return String(a && a.roomTypeName || '').localeCompare(String(b && b.roomTypeName || ''));
}

function validateOtaWorkflowFilters_(input) {
  const payload = input || {};
  const fallbackDate = normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date());
  const startDate = normalizeDateInput_(payload.start_date || payload.startDate || payload.date || fallbackDate) || fallbackDate;
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate || payload.date || startDate) || startDate;
  if (endDate.getTime() < startDate.getTime()) {
    throw new Error('End date must be on or after start date.');
  }

  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || '').trim();
  return {
    selectedDate: normalizeDateInput_(payload.selected_date || payload.selectedDate || startDate) || startDate,
    startDate: startDate,
    endDate: endDate,
    roomTypeId: roomTypeRaw ? resolveRoomTypeId_(roomTypeRaw) : ''
  };
}

function serializeOtaWorkflowFilters_(filters) {
  return {
    selectedDate: formatDateKey_(filters.selectedDate),
    startDate: formatDateKey_(filters.startDate),
    endDate: formatDateKey_(filters.endDate),
    roomTypeId: filters.roomTypeId,
    roomTypeLabel: filters.roomTypeId ? getCompetitorTrackerRoomTypeLabel_(filters.roomTypeId) : '',
    singleDate: formatDateKey_(filters.startDate) === formatDateKey_(filters.endDate)
  };
}

function parseOptionalOtaWorkflowMoney_(value, label) {
  const raw = String(value == null ? '' : value).trim();
  if (raw === '') return null;
  const parsed = Number(raw.replace(/,/g, ''));
  if (!Number.isFinite(parsed) || parsed < 0) {
    throw new Error(label + ' must be a valid amount.');
  }
  return roundCurrency_(parsed);
}

function validateOtaWorkflowEntryInput_(input) {
  const payload = input || {};
  const date = normalizeDateInput_(payload.date || payload.selected_date || payload.selectedDate);
  if (!date) throw new Error('Date is required.');

  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || '').trim();
  if (!roomTypeRaw) throw new Error('Room type is required.');
  const roomTypeId = resolveRoomTypeId_(roomTypeRaw);

  const bookingComStatus = normalizeOtaWorkflowStatus_(payload.booking_com_status || payload.bookingComStatus);
  if (!bookingComStatus || OTA_WORKFLOW_STATUS_VALUES.indexOf(bookingComStatus) === -1) {
    throw new Error('Booking.com status is invalid.');
  }

  const airbnbStatus = normalizeOtaWorkflowStatus_(payload.airbnb_status || payload.airbnbStatus);
  if (!airbnbStatus || OTA_WORKFLOW_STATUS_VALUES.indexOf(airbnbStatus) === -1) {
    throw new Error('Airbnb status is invalid.');
  }

  return {
    date: date,
    roomTypeId: roomTypeId,
    roomTypeName: getRoomTypeNameById_(roomTypeId),
    bookingComStatus: bookingComStatus,
    airbnbStatus: airbnbStatus,
    linkedRecommendationAction: String(payload.linked_recommendation_action || payload.linkedRecommendationAction || '').trim(),
    linkedActionTaken: String(payload.linked_action_taken || payload.linkedActionTaken || '').trim(),
    rozaPriceAfterChange: parseOptionalOtaWorkflowMoney_(payload.roza_price_after_change || payload.rozaPriceAfterChange, 'Roza price after change'),
    note: String(payload.note || '').trim()
  };
}

function buildOtaWorkflowContextKey_(row) {
  return String(row && row.dateKey || '') + '|' + String(row && row.roomTypeId || '');
}

function buildOtaWorkflowSummary_(rows, filters) {
  const latestByContext = {};
  (rows || []).slice().sort(compareOtaWorkflowRows_).forEach(function(row) {
    const key = buildOtaWorkflowContextKey_(row);
    if (!key || latestByContext[key]) return;
    latestByContext[key] = row;
  });
  const latestRows = Object.keys(latestByContext).map(function(key) {
    return latestByContext[key];
  });
  const pendingBookingComCount = latestRows.filter(function(row) {
    return row.bookingComStatus === 'Pending';
  }).length;
  const pendingAirbnbCount = latestRows.filter(function(row) {
    return row.airbnbStatus === 'Pending';
  }).length;
  const pendingAnyCount = latestRows.filter(function(row) {
    return row.bookingComStatus === 'Pending' || row.airbnbStatus === 'Pending';
  }).length;
  const fullyUpdatedCount = latestRows.filter(function(row) {
    return row.bookingComStatus === 'Updated' && row.airbnbStatus === 'Updated';
  }).length;
  const selectedKey = formatDateKey_(filters.selectedDate) + '|' + String(filters.roomTypeId || '');
  const selectedLatestRow = filters.roomTypeId && latestByContext[selectedKey]
    ? serializeOtaWorkflowRow_(latestByContext[selectedKey])
    : null;

  return {
    totalRows: (rows || []).length,
    latestContextCount: latestRows.length,
    pendingBookingComCount: pendingBookingComCount,
    pendingAirbnbCount: pendingAirbnbCount,
    pendingAnyCount: pendingAnyCount,
    fullyUpdatedCount: fullyUpdatedCount,
    selectedLatestRow: selectedLatestRow
  };
}

function buildCompetitorTrackerSummary_(rows, filters) {
  const prices = rows.map(function(row) {
    return Number(row.competitorPrice);
  }).filter(function(value) {
    return Number.isFinite(value);
  });
  const averageCompetitorPrice = prices.length
    ? roundCurrency_(prices.reduce(function(sum, value) { return sum + value; }, 0) / prices.length)
    : null;
  const lowestCompetitorPrice = prices.length ? roundCurrency_(Math.min.apply(null, prices)) : null;
  const highestCompetitorPrice = prices.length ? roundCurrency_(Math.max.apply(null, prices)) : null;
  const competitorSoldOutCount = rows.filter(function(row) {
    return row.availableOrSoldOut === 'Sold Out';
  }).length;
  const totalCompetitorsChecked = rows.length;

  let rozaCurrentPrice = null;
  let rozaPricingSource = '';
  let rozaVsMarketGap = null;
  const isSingleDate = formatDateKey_(filters.startDate) === formatDateKey_(filters.endDate);

  if (filters.roomTypeId && isSingleDate) {
    try {
      const rozaQuote = buildStayAvailabilityPricingSnapshot_(
        filters.startDate,
        addDays_(filters.startDate, 1),
        filters.roomTypeId,
        INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE[filters.roomTypeId] || 2,
        { qtyRooms: 1 }
      );
      if (rozaQuote && rozaQuote.estimatedPrice != null) {
        rozaCurrentPrice = roundCurrency_(rozaQuote.estimatedPrice);
        rozaPricingSource = String(rozaQuote.pricingSource || '').trim();
        if (averageCompetitorPrice != null) {
          rozaVsMarketGap = roundCurrency_(rozaCurrentPrice - averageCompetitorPrice);
        }
      }
    } catch (error) {
      rozaCurrentPrice = null;
      rozaPricingSource = '';
      rozaVsMarketGap = null;
    }
  }

  return {
    startDate: formatDateKey_(filters.startDate),
    endDate: formatDateKey_(filters.endDate),
    roomTypeId: filters.roomTypeId,
    roomTypeLabel: filters.roomTypeId ? getCompetitorTrackerRoomTypeLabel_(filters.roomTypeId) : '',
    singleDate: isSingleDate,
    averageCompetitorPrice: averageCompetitorPrice,
    lowestCompetitorPrice: lowestCompetitorPrice,
    highestCompetitorPrice: highestCompetitorPrice,
    competitorSoldOutCount: competitorSoldOutCount,
    totalCompetitorsChecked: totalCompetitorsChecked,
    priceObservationCount: prices.length,
    rozaCurrentPrice: rozaCurrentPrice,
    rozaPricingSource: rozaPricingSource,
    rozaVsMarketGap: rozaVsMarketGap
  };
}

function validateDemandScoreInput_(input) {
  const payload = input || {};
  const date = normalizeDateInput_(payload.date || payload.selected_date || payload.selectedDate);
  if (!date) throw new Error('Date is required.');
  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || payload.room_type || payload.roomType || '').trim();
  if (!roomTypeRaw) throw new Error('Room type is required.');

  return {
    date: date,
    roomTypeId: resolveRoomTypeId_(roomTypeRaw)
  };
}

function buildPickupPaceData_(targetDate, roomTypeId, options) {
  const settings = options || {};
  const normalizedDate = normalizeDateInput_(targetDate);
  const context = settings.context || buildDemandScoreContext_(roomTypeId);
  const exactDate = settings.exactDate || resolveDemandLiveDate_(normalizedDate, context);
  const asOfDate = stripTime_(settings.asOfDate || new Date());
  const bookingNightRows = settings.bookingNightRows || getSheetObjects_(SHEET_NAMES.BOOKING_NIGHTS);
  const snapshotRows = settings.snapshotRows || getSheetObjects_(SHEET_NAMES.OTB_SNAPSHOTS);
  const currentOtbRooms = Number(exactDate.soldRooms || 0);
  const pickupLast3Days = getPickupPaceBookingWindowRooms_(bookingNightRows, normalizedDate, context.roomTypeId, asOfDate, 3);
  const pickupLast7Days = getPickupPaceBookingWindowRooms_(bookingNightRows, normalizedDate, context.roomTypeId, asOfDate, 7);
  const pickupLast14Days = getPickupPaceBookingWindowRooms_(bookingNightRows, normalizedDate, context.roomTypeId, asOfDate, 14);
  const recentRooms3Days = pickupLast3Days;
  const priorRooms11Days = Math.max(0, pickupLast14Days - pickupLast3Days);
  const snapshotDelta7Days = buildPickupPaceSnapshotSignal_(snapshotRows, asOfDate, normalizedDate, context.roomTypeId, 7);
  const snapshotDelta14Days = buildPickupPaceSnapshotSignal_(snapshotRows, asOfDate, normalizedDate, context.roomTypeId, 14);
  const snapshotOtbRooms7DaysAgo = snapshotDelta7Days.available ? Number(snapshotDelta7Days.previousRoomsSold || 0) : null;
  const snapshotOtbRooms14DaysAgo = snapshotDelta14Days.available ? Number(snapshotDelta14Days.previousRoomsSold || 0) : null;
  const paceDelta7Days = snapshotOtbRooms7DaysAgo == null ? null : currentOtbRooms - snapshotOtbRooms7DaysAgo;
  const paceDelta14Days = snapshotOtbRooms14DaysAgo == null ? null : currentOtbRooms - snapshotOtbRooms14DaysAgo;
  const bookingHistoryAvailable = Array.isArray(bookingNightRows);
  const snapshotHistoryAvailable = !!(snapshotDelta7Days.available || snapshotDelta14Days.available);
  const pickupSignalRooms = snapshotHistoryAvailable
    ? paceDelta7Days
    : (bookingHistoryAvailable ? pickupLast7Days : null);
  const pickupBand = pickupSignalRooms != null
    ? getPickupPaceBand_(pickupSignalRooms, Number(exactDate.sellableInventory || 0))
    : '';
  const pickupTrend = bookingHistoryAvailable
    ? getPickupPaceTrend_(recentRooms3Days, priorRooms11Days)
    : getPickupPaceSnapshotTrend_(paceDelta7Days, paceDelta14Days);

  return {
    ok: true,
    available: bookingHistoryAvailable || snapshotHistoryAvailable,
    date: formatDateKey_(normalizedDate),
    roomTypeId: context.roomTypeId,
    roomTypeLabel: context.roomTypeLabel,
    currency: getDefaultBookingCurrency_(),
    asOfDate: formatDateKey_(asOfDate),
    currentOtbRooms: currentOtbRooms,
    current_otb_rooms: currentOtbRooms,
    pickupLast3Days: pickupLast3Days,
    pickup_last_3_days: pickupLast3Days,
    pickupLast7Days: pickupLast7Days,
    pickup_last_7_days: pickupLast7Days,
    pickupLast14Days: pickupLast14Days,
    pickup_last_14_days: pickupLast14Days,
    pickupBand: pickupBand,
    pickup_band: pickupBand,
    pickupTrend: pickupTrend,
    pickup_trend: pickupTrend,
    currentSnapshotDate: snapshotDelta7Days.currentSnapshotDate || snapshotDelta14Days.currentSnapshotDate || '',
    current_snapshot_date: snapshotDelta7Days.currentSnapshotDate || snapshotDelta14Days.currentSnapshotDate || '',
    snapshotOtbRooms7DaysAgo: snapshotOtbRooms7DaysAgo,
    snapshot_otb_rooms_7_days_ago: snapshotOtbRooms7DaysAgo,
    snapshotOtbRooms14DaysAgo: snapshotOtbRooms14DaysAgo,
    snapshot_otb_rooms_14_days_ago: snapshotOtbRooms14DaysAgo,
    paceDelta7Days: paceDelta7Days,
    pace_delta_7_days: paceDelta7Days,
    paceDelta14Days: paceDelta14Days,
    pace_delta_14_days: paceDelta14Days,
    bookingHistoryAvailable: bookingHistoryAvailable,
    snapshotHistoryAvailable: snapshotHistoryAvailable,
    pickupSignalRooms: pickupSignalRooms,
    pickup_signal_rooms: pickupSignalRooms,
    signalSource: paceDelta7Days != null ? 'snapshot_delta_7_day' : (bookingHistoryAvailable ? 'booking_created_history_7_day' : ''),
    signal_source: paceDelta7Days != null ? 'snapshot_delta_7_day' : (bookingHistoryAvailable ? 'booking_created_history_7_day' : ''),
    recentRooms3Days: recentRooms3Days,
    priorRooms11Days: priorRooms11Days,
    bookingHistoryNote: bookingHistoryAvailable
      ? 'Pickup counts use current Booking_Nights created-at history for this stay date and room type.'
      : 'Booking created-at history is not available for this stay date yet.',
    snapshotHistoryNote: buildPickupPaceSnapshotNote_(snapshotDelta7Days, snapshotDelta14Days)
  };
}

function buildDemandScoreData_(targetDate, roomTypeId) {
  const normalizedDate = normalizeDateInput_(targetDate);
  const context = buildDemandScoreContext_(roomTypeId);
  const exactDate = resolveDemandLiveDate_(normalizedDate, context);
  const occupancyPct = safeDivide_(Number(exactDate.soldRooms || 0), Number(exactDate.sellableInventory || 0));
  const competitorSummary = getCompetitorTrackerData({
    start_date: normalizedDate,
    end_date: normalizedDate,
    room_type_id: context.roomTypeId
  }).summary || {};
  const eventFlagSummary = buildEventFlagSelectedDateSummary_(
    normalizedDate,
    getSheetObjects_(SHEET_NAMES.EVENT_FLAGS).map(function(row, index) {
      return normalizeEventFlagRow_(row, index + 2);
    })
  );
  const snapshotRows = getSheetObjects_(SHEET_NAMES.OTB_SNAPSHOTS);
  const bookingNightRows = getSheetObjects_(SHEET_NAMES.BOOKING_NIGHTS);
  const pickupMetrics = buildPickupPaceData_(normalizedDate, context.roomTypeId, {
    context: context,
    exactDate: exactDate,
    snapshotRows: snapshotRows,
    bookingNightRows: bookingNightRows
  });
  const pressure7 = buildDemandPressureWindow_(normalizedDate, context, 7);
  const pressure30 = buildDemandPressureWindow_(normalizedDate, context, 30);
  const pressure90 = buildDemandPressureWindow_(normalizedDate, context, 90);
  const scoreResult = calculateDemandScore_(exactDate, occupancyPct, competitorSummary, pickupMetrics, pressure7, eventFlagSummary);
  const currency = getDefaultBookingCurrency_();

  return {
    ok: true,
    date: formatDateKey_(normalizedDate),
    roomTypeId: context.roomTypeId,
    roomTypeLabel: context.roomTypeLabel,
    currency: currency,
    demandScore: scoreResult.score,
    demandBand: getDemandBandFromScore_(scoreResult.score),
    usedSignals: scoreResult.usedSignals,
    componentScores: scoreResult.componentScores,
    exactDateMetrics: {
      occupancyPct: occupancyPct,
      soldRooms: Number(exactDate.soldRooms || 0),
      sellableInventory: Number(exactDate.sellableInventory || 0),
      physicalInventory: Number(exactDate.physicalInventory || 0),
      blockedRooms: Number(exactDate.blockedRooms || 0),
      availableRooms: Number(exactDate.availableRooms || 0),
      overbookingAllowanceApplied: Number(exactDate.overbookingAllowanceApplied || 0)
    },
    competitorMetrics: {
      competitorSoldOutCount: Number(competitorSummary.competitorSoldOutCount || 0),
      totalCompetitorsChecked: Number(competitorSummary.totalCompetitorsChecked || 0),
      competitorAveragePrice: competitorSummary.averageCompetitorPrice == null ? null : Number(competitorSummary.averageCompetitorPrice),
      competitorLowestPrice: competitorSummary.lowestCompetitorPrice == null ? null : Number(competitorSummary.lowestCompetitorPrice),
      rozaCurrentPrice: competitorSummary.rozaCurrentPrice == null ? null : Number(competitorSummary.rozaCurrentPrice),
      rozaVsMarketGap: competitorSummary.rozaVsMarketGap == null ? null : Number(competitorSummary.rozaVsMarketGap),
      rozaPricingSource: String(competitorSummary.rozaPricingSource || '').trim()
    },
    eventFlagSummary: eventFlagSummary,
    pickupMetrics: pickupMetrics,
    next7DayPressure: pressure7,
    next30DayPressure: pressure30,
    next90DayPressure: pressure90,
    manualEventFlagUsed: scoreResult.usedSignals.indexOf('event_flag') !== -1
  };
}

function buildPricingRecommendationData_(targetDate, roomTypeId) {
  const demandData = buildDemandScoreData_(targetDate, roomTypeId);
  const exact = demandData.exactDateMetrics || {};
  const competitor = demandData.competitorMetrics || {};
  const recommendation = calculatePricingRecommendation_(demandData);
  const occupancyPct = Number(exact.occupancyPct || 0);
  const overbookingAllowanceUnits = Number(exact.overbookingAllowanceApplied || 0);
  const rozaCurrentPrice = competitor.rozaCurrentPrice == null ? null : Number(competitor.rozaCurrentPrice);
  const pricingSource = String(competitor.rozaPricingSource || '').trim();
  const competitorAveragePrice = competitor.competitorAveragePrice == null ? null : Number(competitor.competitorAveragePrice);
  const competitorLowestPrice = competitor.competitorLowestPrice == null ? null : Number(competitor.competitorLowestPrice);
  const rozaVsMarketGap = competitor.rozaVsMarketGap == null ? null : Number(competitor.rozaVsMarketGap);
  const competitorSoldOutCount = Number(competitor.competitorSoldOutCount || 0);
  const totalCompetitorsChecked = Number(competitor.totalCompetitorsChecked || 0);
  const eventFlagSummary = demandData.eventFlagSummary || {};

  return {
    ok: true,
    date: demandData.date,
    roomTypeId: demandData.roomTypeId,
    roomTypeLabel: demandData.roomTypeLabel,
    currency: demandData.currency,
    recommendationAction: recommendation.action,
    recommendationReasonSummary: recommendation.reasonSummary,
    recommendationConfidence: recommendation.confidence,
    recommendationSignalsUsed: recommendation.signalsUsed,
    recommendationNotes: recommendation.notes,
    recommendation_action: recommendation.action,
    recommendation_reason_summary: recommendation.reasonSummary,
    recommendation_confidence: recommendation.confidence,
    rozaCurrentPrice: rozaCurrentPrice,
    roza_current_price: rozaCurrentPrice,
    pricingSource: pricingSource,
    pricing_source: pricingSource,
    competitorAveragePrice: competitorAveragePrice,
    competitor_average_price: competitorAveragePrice,
    competitorLowestPrice: competitorLowestPrice,
    competitor_lowest_price: competitorLowestPrice,
    rozaVsMarketGap: rozaVsMarketGap,
    roza_vs_market_gap: rozaVsMarketGap,
    demandScore: Number(demandData.demandScore || 0),
    demand_score: Number(demandData.demandScore || 0),
    demandBand: String(demandData.demandBand || 'Low'),
    demand_band: String(demandData.demandBand || 'Low'),
    next7DayPressure: demandData.next7DayPressure || null,
    next30DayPressure: demandData.next30DayPressure || null,
    next90DayPressure: demandData.next90DayPressure || null,
    next_7_day_pressure: demandData.next7DayPressure ? String(demandData.next7DayPressure.band || 'Low') : 'Low',
    next_30_day_pressure: demandData.next30DayPressure ? String(demandData.next30DayPressure.band || 'Low') : 'Low',
    next_90_day_pressure: demandData.next90DayPressure ? String(demandData.next90DayPressure.band || 'Low') : 'Low',
    competitorSoldOutCount: competitorSoldOutCount,
    competitor_sold_out_count: competitorSoldOutCount,
    totalCompetitorsChecked: totalCompetitorsChecked,
    occupancyPct: occupancyPct,
    occupancy_pct: occupancyPct,
    overbookingAllowanceActive: overbookingAllowanceUnits > 0,
    overbooking_allowance_active: overbookingAllowanceUnits > 0,
    overbookingAllowanceUnits: overbookingAllowanceUnits,
    pickupBand: demandData.pickupMetrics && demandData.pickupMetrics.pickupBand ? String(demandData.pickupMetrics.pickupBand) : '',
    pickup_band: demandData.pickupMetrics && demandData.pickupMetrics.pickupBand ? String(demandData.pickupMetrics.pickupBand) : '',
    pickupTrend: demandData.pickupMetrics && demandData.pickupMetrics.pickupTrend ? String(demandData.pickupMetrics.pickupTrend) : '',
    pickup_trend: demandData.pickupMetrics && demandData.pickupMetrics.pickupTrend ? String(demandData.pickupMetrics.pickupTrend) : '',
    exactDateMetrics: exact,
    competitorMetrics: competitor,
    eventFlagSummary: eventFlagSummary,
    pickupMetrics: demandData.pickupMetrics || {},
    usedSignals: demandData.usedSignals || [],
    componentScores: demandData.componentScores || {},
    manualEventFlagUsed: !!demandData.manualEventFlagUsed
  };
}

function applyManualCommercialShortcut_(validated) {
  const targetDate = normalizeDateInput_(validated.date);
  const roomTypeId = resolveRoomTypeId_(validated.roomTypeId);
  const guests = Number(INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE[roomTypeId] || 2);
  const before = resolveLiveCommercialDate_(targetDate, roomTypeId, guests);
  const currentLivePrice = before && before.finalNightlyRate != null ? roundCurrency_(before.finalNightlyRate) : null;
  const actionType = normalizeManualCommercialShortcutAction_(validated.action);

  if (currentLivePrice == null || Number(currentLivePrice) <= 0) {
    throw new Error('A live price greater than zero is required before a manual shortcut can be applied.');
  }

  let appliedDelta = null;
  let exactTarget = null;
  let nextPrice = currentLivePrice;
  if (actionType === 'exact_price') {
    exactTarget = roundCurrency_(validated.targetPrice);
    nextPrice = exactTarget;
    appliedDelta = roundCurrency_(nextPrice - currentLivePrice);
  } else {
    appliedDelta = actionType === 'custom_delta'
      ? roundCurrency_(validated.customDelta)
      : getManualCommercialShortcutDelta_(validated.action);
    nextPrice = roundCurrency_(actionType === 'hold' ? currentLivePrice : currentLivePrice + appliedDelta);
  }

  if (!Number.isFinite(Number(nextPrice)) || Number(nextPrice) <= 0) {
    throw new Error('This shortcut would produce an invalid live price.');
  }

  const controls = getCommercialControlRows_({ activeOnly: false, roomTypeId: roomTypeId });
  const exactSpecialControl = getExactSpecialCommercialControlForDate_(roomTypeId, targetDate, controls);
  const saveResult = adminSaveCommercialControl({
    control_id: exactSpecialControl ? exactSpecialControl.controlId : '',
    room_type_id: roomTypeId,
    rule_type: 'special',
    start_date: targetDate,
    end_date: targetDate,
    override_price: nextPrice,
    overbooking_allowance: exactSpecialControl ? Number(exactSpecialControl.overbookingAllowance || 0) : 0,
    active: 'Yes',
    note: buildManualCommercialShortcutNote_(validated.action, currentLivePrice, nextPrice, validated.operatorNote, validated.customDelta, validated.targetPrice),
    selected_date: validated.selectedDate || targetDate
  });
  const after = resolveLiveCommercialDate_(targetDate, roomTypeId, guests);
  const roomTypeLabel = getCompetitorTrackerRoomTypeLabel_(roomTypeId);
  const actionLabel = getManualCommercialShortcutActionLabel_(validated.action, validated.customDelta, validated.targetPrice);
  const loggedActionLabel = getManualCommercialShortcutLoggedAction_(validated.action, validated.customDelta, validated.targetPrice);
  const currency = getDefaultBookingCurrency_();
  const newLivePrice = after && after.finalNightlyRate != null ? roundCurrency_(after.finalNightlyRate) : nextPrice;
  let actionLogEntry = null;
  let actionLogData = null;

  if (validated.recommendationAction) {
    const logResult = adminSaveRecommendationActionLog({
      date: targetDate,
      room_type_id: roomTypeId,
      recommendation_action: validated.recommendationAction,
      recommendation_confidence: validated.recommendationConfidence,
      recommendation_reason_summary: validated.recommendationReasonSummary,
      roza_old_price: currentLivePrice,
      action_taken: loggedActionLabel,
      roza_new_price: newLivePrice,
      overbooking_change: 'No change',
      ota_updated: 'No',
      operator_note: validated.operatorNote
    });
    actionLogEntry = logResult.entry || null;
    actionLogData = logResult.logData || null;
  }

  return {
    ok: true,
    action: validated.action,
    actionType: actionType,
    actionLabel: actionLabel,
    loggedActionLabel: loggedActionLabel,
    date: formatDateKey_(targetDate),
    roomTypeId: roomTypeId,
    roomTypeLabel: roomTypeLabel,
    currency: currency,
    controlId: saveResult.controlId,
    oldLivePrice: currentLivePrice,
    newLivePrice: newLivePrice,
    appliedDelta: appliedDelta,
    exactTarget: exactTarget,
    oldPricingSource: String(before && before.pricingSource || '').trim(),
    newPricingSource: String(after && after.pricingSource || '').trim(),
    actionLogEntry: actionLogEntry,
    actionLogData: actionLogData,
    updatedAt: new Date(),
    message: buildManualCommercialShortcutResultMessage_(actionLabel, currentLivePrice, newLivePrice, currency),
    controls: saveResult.controls
  };
}

function calculatePricingRecommendation_(demandData) {
  const exact = demandData && demandData.exactDateMetrics ? demandData.exactDateMetrics : {};
  const competitor = demandData && demandData.competitorMetrics ? demandData.competitorMetrics : {};
  const pickup = demandData && demandData.pickupMetrics ? demandData.pickupMetrics : {};
  const pressure7 = demandData && demandData.next7DayPressure ? demandData.next7DayPressure : {};
  const pressure30 = demandData && demandData.next30DayPressure ? demandData.next30DayPressure : {};
  const eventFlagSummary = demandData && demandData.eventFlagSummary ? demandData.eventFlagSummary : {};
  const demandScore = Number(demandData && demandData.demandScore || 0);
  const demandBand = String(demandData && demandData.demandBand || 'Low');
  const occupancyPct = Number(exact.occupancyPct || 0);
  const availableRooms = Number(exact.availableRooms || 0);
  const sellableInventory = Number(exact.sellableInventory || 0);
  const overbookingAllowanceUnits = Number(exact.overbookingAllowanceApplied || 0);
  const rozaCurrentPrice = competitor.rozaCurrentPrice == null ? null : Number(competitor.rozaCurrentPrice);
  const competitorAveragePrice = competitor.competitorAveragePrice == null ? null : Number(competitor.competitorAveragePrice);
  const competitorLowestPrice = competitor.competitorLowestPrice == null ? null : Number(competitor.competitorLowestPrice);
  const rozaVsMarketGap = competitor.rozaVsMarketGap == null ? null : Number(competitor.rozaVsMarketGap);
  const competitorSoldOutCount = Number(competitor.competitorSoldOutCount || 0);
  const totalCompetitorsChecked = Number(competitor.totalCompetitorsChecked || 0);
  const soldOutRatio = totalCompetitorsChecked > 0 ? safeDivide_(competitorSoldOutCount, totalCompetitorsChecked) : 0;
  const pickupSignalRooms = pickup && pickup.pickupSignalRooms != null ? Number(pickup.pickupSignalRooms) : null;
  const pickupBand = String(pickup && pickup.pickupBand || '');
  const pickupTrend = String(pickup && pickup.pickupTrend || '');
  const positiveSignals = [];
  const negativeSignals = [];
  const missingSignals = [];
  let raiseStrength = 0;
  let dropStrength = 0;
  let action = 'Hold';
  let manualReason = '';

  if (rozaCurrentPrice == null) {
    missingSignals.push('Live Roza price unavailable');
  }
  if (competitorAveragePrice == null) {
    missingSignals.push('Competitor average unavailable');
  }
  if (!pickup || !pickup.available) {
    missingSignals.push('Pickup unavailable');
  }
  if (sellableInventory <= 0) {
    missingSignals.push('Sellable inventory unavailable');
  }

  if (demandBand === 'Hot') {
    raiseStrength += 3;
    positiveSignals.push('Demand is hot');
  } else if (demandBand === 'Strong') {
    raiseStrength += 2;
    positiveSignals.push('Demand is strong');
  } else if (demandBand === 'Low') {
    dropStrength += 2;
    negativeSignals.push('Demand is low');
  }

  if (eventFlagSummary && eventFlagSummary.hasActiveFlag) {
    if (Number(eventFlagSummary.highestImpactRank || 0) >= 3) {
      raiseStrength += 1;
      positiveSignals.push('high-impact event flag is active');
    } else {
      positiveSignals.push('event flag is active');
    }
  }

  if (occupancyPct >= 0.95 || availableRooms <= 0) {
    raiseStrength += 2;
    positiveSignals.push('inventory is fully compressed');
  } else if (occupancyPct >= 0.85 || availableRooms <= 1) {
    raiseStrength += 1;
    positiveSignals.push('inventory is tight');
  } else if (occupancyPct < 0.30) {
    dropStrength += 2;
    negativeSignals.push('occupancy is weak');
  } else if (occupancyPct < 0.45) {
    dropStrength += 1;
    negativeSignals.push('occupancy is soft');
  }

  if (getDemandBandRank_(pressure7.band) >= 3) {
    raiseStrength += 1;
    positiveSignals.push('next 7 days are under pressure');
  } else if (String(pressure7.band || '') === 'Low') {
    dropStrength += 1;
    negativeSignals.push('next 7 days look soft');
  }

  if (getDemandBandRank_(pressure30.band) >= 3) {
    raiseStrength += 1;
    positiveSignals.push('next 30 days stay firm');
  } else if (String(pressure30.band || '') === 'Low') {
    dropStrength += 1;
    negativeSignals.push('next 30 days remain soft');
  }

  if (totalCompetitorsChecked > 0) {
    if (soldOutRatio >= 0.75) {
      raiseStrength += 2;
      positiveSignals.push('most checked competitors are sold out');
    } else if (soldOutRatio >= 0.40) {
      raiseStrength += 1;
      positiveSignals.push('competitor sell-out pressure is rising');
    } else if (soldOutRatio === 0 && totalCompetitorsChecked >= 2) {
      dropStrength += 1;
      negativeSignals.push('competitors still have space');
    }
  } else {
    missingSignals.push('No competitor checks loaded');
  }

  if (rozaVsMarketGap != null) {
    if (rozaVsMarketGap <= -15) {
      raiseStrength += 3;
      positiveSignals.push('Roza sits well below market average');
    } else if (rozaVsMarketGap <= -8) {
      raiseStrength += 2;
      positiveSignals.push('Roza sits below market average');
    } else if (rozaVsMarketGap <= -3) {
      raiseStrength += 1;
      positiveSignals.push('Roza is slightly below market');
    } else if (rozaVsMarketGap >= 10) {
      dropStrength += 2;
      negativeSignals.push('Roza is above market average');
    } else if (rozaVsMarketGap >= 5) {
      dropStrength += 1;
      negativeSignals.push('Roza is slightly above market');
    }
  }

  if (competitorLowestPrice != null && rozaCurrentPrice != null) {
    if (rozaCurrentPrice <= competitorLowestPrice - 8) {
      raiseStrength += 1;
      positiveSignals.push('Roza is below the cheapest checked competitor');
    } else if (rozaCurrentPrice >= competitorLowestPrice + 10) {
      dropStrength += 1;
      negativeSignals.push('Roza is above the cheapest checked competitor');
    }
  }

  if (pickupSignalRooms != null) {
    if (pickupSignalRooms >= 2) {
      raiseStrength += 1;
      positiveSignals.push('pickup added rooms over the last 7 days');
    } else if (pickupSignalRooms < 0) {
      dropStrength += 1;
      negativeSignals.push('pickup softened versus the last 7 days');
    }
  }

  if (pickupBand === 'Strong') {
    raiseStrength += 1;
    positiveSignals.push('pickup is strong');
  } else if (pickupBand === 'Weak') {
    dropStrength += 1;
    negativeSignals.push('pickup is weak');
  }

  if (pickupTrend === 'Improving') {
    raiseStrength += 1;
    positiveSignals.push('pickup pace is improving');
  } else if (pickupTrend === 'Slowing') {
    dropStrength += 1;
    negativeSignals.push('pickup pace is slowing');
  }

  if (rozaCurrentPrice == null || sellableInventory <= 0) {
    action = 'Review manually';
    manualReason = 'missing_data';
  } else if (demandScore >= 80 && availableRooms <= 0 && soldOutRatio >= 0.50 && overbookingAllowanceUnits <= 0) {
    action = 'Review manually';
    manualReason = 'compression_review';
  } else if (raiseStrength >= 8 && dropStrength <= 1) {
    action = 'Raise ' + formatManualShortcutPoundsCompact_(20);
  } else if (raiseStrength >= 6 && dropStrength <= 2) {
    action = 'Raise ' + formatManualShortcutPoundsCompact_(10);
  } else if (raiseStrength >= 4 && dropStrength <= 2) {
    action = 'Raise ' + formatManualShortcutPoundsCompact_(5);
  } else if (dropStrength >= 6 && raiseStrength <= 1) {
    action = 'Drop ' + formatManualShortcutPoundsCompact_(10);
  } else if (dropStrength >= 4 && raiseStrength <= 2) {
    action = 'Drop ' + formatManualShortcutPoundsCompact_(5);
  } else if (Math.abs(raiseStrength - dropStrength) <= 1 && rozaVsMarketGap != null && Math.abs(rozaVsMarketGap) < 5) {
    action = 'Hold';
  } else if (demandBand === 'Normal' && (!rozaVsMarketGap || Math.abs(rozaVsMarketGap) < 8)) {
    action = 'Hold';
  } else {
    action = 'Review manually';
    manualReason = 'mixed_signals';
  }

  const hasConflict = raiseStrength >= 3 && dropStrength >= 3;
  const signalCount = countRecommendationSignals_(rozaCurrentPrice, competitorAveragePrice, totalCompetitorsChecked, pickupSignalRooms, pressure7, pressure30, sellableInventory, eventFlagSummary);
  const confidence = getPricingRecommendationConfidence_(action, signalCount, hasConflict, competitorAveragePrice != null, totalCompetitorsChecked, manualReason);

  return {
    action: action,
    reasonSummary: buildPricingRecommendationReasonSummary_(action, positiveSignals, negativeSignals, manualReason),
    confidence: confidence,
    signalsUsed: dedupeRecommendationNotes_(positiveSignals.concat(negativeSignals)),
    notes: {
      raiseStrength: raiseStrength,
      dropStrength: dropStrength,
      missingSignals: dedupeRecommendationNotes_(missingSignals),
      manualReason: manualReason
    }
  };
}

function buildPricingRecommendationReasonSummary_(action, positiveSignals, negativeSignals, manualReason) {
  if (action === 'Review manually') {
    if (manualReason === 'compression_review') {
      return 'Hot compression and strong market sell-out are visible. Review manually before using any overbooking move.';
    }
    if (manualReason === 'missing_data') {
      return 'Key pricing signals are incomplete, so this date needs manual commercial review.';
    }
    return 'Signals are mixed or not strong enough for a clean rule-based move. Review manually.';
  }

  if (action === 'Hold') {
    return 'Demand and market signals are broadly balanced around the current live price.';
  }

  const notes = dedupeRecommendationNotes_(action.indexOf('Raise') === 0 ? positiveSignals : negativeSignals);
  const selected = notes.slice(0, 3);
  if (!selected.length) {
    return action.indexOf('Raise') === 0
      ? 'Demand looks firmer than the current live price suggests.'
      : 'Demand looks softer than the current live price suggests.';
  }
  return capitalizeRecommendationSentence_(selected.join(', ')) + '.';
}

function validateManualCommercialShortcutInput_(input) {
  const payload = input || {};
  const demandInput = validateDemandScoreInput_(payload);
  const action = normalizeManualCommercialShortcutAction_(payload.action || payload.shortcut_action || payload.shortcutAction);
  const operatorNote = String(payload.operator_note || payload.operatorNote || '').trim();
  if (!action) {
    throw new Error('Shortcut action is required.');
  }
  const customDeltaRaw = String(payload.custom_delta == null ? '' : payload.custom_delta).trim();
  const targetPriceRaw = String(payload.target_price == null ? '' : payload.target_price).trim();
  let customDelta = null;
  let targetPrice = null;

  if ((action === 'custom_delta' || action === 'exact_price') && customDeltaRaw && targetPriceRaw) {
    throw new Error('Use either a custom adjustment or an exact target price, not both.');
  }
  if (action === 'custom_delta') {
    if (!customDeltaRaw) {
      throw new Error('Enter a custom increase or decrease amount before applying.');
    }
    customDelta = parseManualCommercialShortcutAmount_(customDeltaRaw, 'Custom adjustment');
    if (Number(customDelta) === 0) {
      throw new Error('Custom adjustment must be greater than zero or less than zero.');
    }
  }
  if (action === 'exact_price') {
    if (!targetPriceRaw) {
      throw new Error('Enter an exact target price before applying.');
    }
    targetPrice = parseManualCommercialShortcutAmount_(targetPriceRaw, 'Exact target price');
    if (Number(targetPrice) <= 0) {
      throw new Error('Exact target price must be greater than zero.');
    }
  }
  if (action !== 'hold' && !operatorNote) {
    throw new Error('Add an operator note before applying a manual price change.');
  }

  return {
    date: demandInput.date,
    roomTypeId: demandInput.roomTypeId,
    action: action,
    customDelta: customDelta,
    targetPrice: targetPrice,
    operatorNote: operatorNote,
    recommendationAction: String(payload.recommendation_action || payload.recommendationAction || '').trim(),
    recommendationConfidence: String(payload.recommendation_confidence || payload.recommendationConfidence || '').trim(),
    recommendationReasonSummary: String(payload.recommendation_reason_summary || payload.recommendationReasonSummary || '').trim(),
    selectedDate: normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || payload.date || new Date())
  };
}

function normalizeManualCommercialShortcutAction_(value) {
  const normalized = String(value || '').trim().toLowerCase();
  const allowed = {
    hold: true,
    raise_5: true,
    raise_10: true,
    raise_20: true,
    drop_5: true,
    drop_10: true,
    custom_delta: true,
    exact_price: true
  };
  return allowed[normalized] ? normalized : '';
}

function parseManualCommercialShortcutAmount_(value, label) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) {
    throw new Error(label + ' is required.');
  }
  const parsed = Number(raw.replace(/,/g, ''));
  if (!Number.isFinite(parsed)) {
    throw new Error(label + ' must be a valid amount.');
  }
  return roundCurrency_(parsed);
}

function getManualCommercialShortcutDelta_(action) {
  const normalized = normalizeManualCommercialShortcutAction_(action);
  if (normalized === 'raise_5') return 5;
  if (normalized === 'raise_10') return 10;
  if (normalized === 'raise_20') return 20;
  if (normalized === 'drop_5') return -5;
  if (normalized === 'drop_10') return -10;
  return 0;
}

function formatManualShortcutPoundsCompact_(value) {
  const amount = Math.abs(roundCurrency_(Number(value || 0)));
  const currency = getDefaultBookingCurrency_();
  if (!Number.isFinite(amount)) return currency + ' 0';
  const text = Math.abs(amount - Math.round(amount)) < 0.001
    ? String(Math.round(amount))
    : amount.toFixed(2).replace(/\.?0+$/, '');
  return currency + ' ' + text;
}

function getManualCommercialShortcutActionLabel_(action, customDelta, targetPrice) {
  const normalized = normalizeManualCommercialShortcutAction_(action);
  if (normalized === 'raise_5') return 'Apply +' + formatManualShortcutPoundsCompact_(5);
  if (normalized === 'raise_10') return 'Apply +' + formatManualShortcutPoundsCompact_(10);
  if (normalized === 'raise_20') return 'Apply +' + formatManualShortcutPoundsCompact_(20);
  if (normalized === 'drop_5') return 'Apply -' + formatManualShortcutPoundsCompact_(5);
  if (normalized === 'drop_10') return 'Apply -' + formatManualShortcutPoundsCompact_(10);
  if (normalized === 'custom_delta') {
    const amount = Number(customDelta || 0);
    return 'Apply custom ' + (amount >= 0 ? '+' : '-') + formatManualShortcutPoundsCompact_(amount);
  }
  if (normalized === 'exact_price') {
    return 'Set exact price ' + formatManualShortcutPoundsCompact_(targetPrice);
  }
  return 'Hold';
}

function getManualCommercialShortcutLoggedAction_(action, customDelta, targetPrice) {
  const normalized = normalizeManualCommercialShortcutAction_(action);
  if (normalized === 'raise_5') return 'Raised ' + formatManualShortcutPoundsCompact_(5);
  if (normalized === 'raise_10') return 'Raised ' + formatManualShortcutPoundsCompact_(10);
  if (normalized === 'raise_20') return 'Raised ' + formatManualShortcutPoundsCompact_(20);
  if (normalized === 'drop_5') return 'Dropped ' + formatManualShortcutPoundsCompact_(5);
  if (normalized === 'drop_10') return 'Dropped ' + formatManualShortcutPoundsCompact_(10);
  if (normalized === 'custom_delta') {
    return Number(customDelta || 0) >= 0 ? 'Raised custom' : 'Dropped custom';
  }
  if (normalized === 'exact_price') {
    return 'Set exact price';
  }
  return 'Held price';
}

function buildManualCommercialShortcutNote_(action, oldPrice, newPrice, operatorNote, customDelta, targetPrice) {
  const parts = [
    'Manual apply shortcut',
    getManualCommercialShortcutActionLabel_(action, customDelta, targetPrice),
    'from ' + formatManualShortcutMoney_(oldPrice),
    'to ' + formatManualShortcutMoney_(newPrice)
  ];
  if (operatorNote) parts.push(operatorNote);
  return parts.join(' | ');
}

function buildManualCommercialShortcutResultMessage_(actionLabel, oldPrice, newPrice, currency) {
  return actionLabel + ' applied. ' +
    formatManualShortcutMoney_(oldPrice, currency) + ' -> ' +
    formatManualShortcutMoney_(newPrice, currency) + '.';
}

function formatManualShortcutMoney_(value, currency) {
  const amount = Number(value || 0);
  const code = String(currency || getDefaultBookingCurrency_() || 'GBP').trim().toUpperCase();
  if (!Number.isFinite(amount)) return code + ' -';
  return code + ' ' + amount.toFixed(2);
}

function getPricingRecommendationConfidence_(action, signalCount, hasConflict, hasCompetitorPrice, totalCompetitorsChecked, manualReason) {
  if (hasConflict) return 'Low';
  if (action === 'Review manually') {
    return manualReason === 'compression_review' && signalCount >= 5 ? 'Medium' : 'Low';
  }
  if (signalCount >= 6 && hasCompetitorPrice && Number(totalCompetitorsChecked || 0) >= 2) return 'High';
  if (signalCount >= 4) return 'Medium';
  return 'Low';
}

function countRecommendationSignals_(rozaCurrentPrice, competitorAveragePrice, totalCompetitorsChecked, pickupRoomsDelta, pressure7, pressure30, sellableInventory, eventFlagSummary) {
  let signalCount = 0;
  if (sellableInventory > 0) signalCount += 1;
  if (rozaCurrentPrice != null) signalCount += 1;
  if (competitorAveragePrice != null) signalCount += 1;
  if (Number(totalCompetitorsChecked || 0) > 0) signalCount += 1;
  if (pickupRoomsDelta != null) signalCount += 1;
  if (pressure7 && pressure7.band) signalCount += 1;
  if (pressure30 && pressure30.band) signalCount += 1;
  if (eventFlagSummary && eventFlagSummary.hasActiveFlag) signalCount += 1;
  return signalCount;
}

function dedupeRecommendationNotes_(notes) {
  const seen = {};
  return (notes || []).filter(function(note) {
    const key = String(note || '').trim().toLowerCase();
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function capitalizeRecommendationSentence_(text) {
  const value = String(text || '').trim();
  return value ? value.charAt(0).toUpperCase() + value.slice(1) : '';
}

function buildDemandScoreContext_(roomTypeId) {
  const resolvedRoomTypeId = resolveRoomTypeId_(roomTypeId);
  return {
    roomTypeId: resolvedRoomTypeId,
    roomTypeLabel: getCompetitorTrackerRoomTypeLabel_(resolvedRoomTypeId),
    guests: Number(INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE[resolvedRoomTypeId] || 2),
    physicalInventory: getRoomInventory_(resolvedRoomTypeId),
    commercialControls: getCommercialControlRows_({ activeOnly: true, roomTypeId: resolvedRoomTypeId }),
    baseRateRows: getBaseRateRows_({ activeOnly: true, roomTypeId: resolvedRoomTypeId }),
    legacyRateRows: getOpenRateRowsForType_(resolvedRoomTypeId),
    bookingCountMap: buildConfirmedBookingRoomCountMap_(getSheetObjects_(SHEET_NAMES.BOOKINGS)).byDateRoom,
    blockedCountMap: buildBlockedDateQtyMap_(getSheetObjects_(SHEET_NAMES.BLOCKED_DATES)).byDateRoom
  };
}

function resolveDemandLiveDate_(date, context) {
  return resolveLiveCommercialDate_(date, context.roomTypeId, context.guests, {
    physicalInventory: context.physicalInventory,
    commercialControls: context.commercialControls,
    baseRateRows: context.baseRateRows,
    legacyRateRows: context.legacyRateRows,
    bookingCountMap: context.bookingCountMap,
    blockedCountMap: context.blockedCountMap
  });
}

function buildDemandPressureWindow_(targetDate, context, days) {
  const windowDays = Math.max(1, Number(days || 1));
  let occupancySum = 0;
  let soldOutDays = 0;
  let compressionDays = 0;
  let overbookingDays = 0;

  for (let index = 0; index < windowDays; index++) {
    const currentDate = addDays_(targetDate, index);
    const metrics = resolveDemandLiveDate_(currentDate, context);
    const occupancyPct = safeDivide_(Number(metrics.soldRooms || 0), Number(metrics.sellableInventory || 0));
    occupancySum += occupancyPct;
    if (Number(metrics.availableRooms || 0) <= 0) soldOutDays += 1;
    if (Number(metrics.availableRooms || 0) <= 1) compressionDays += 1;
    if (Number(metrics.overbookingAllowanceApplied || 0) > 0) overbookingDays += 1;
  }

  const averageOccupancyPct = safeDivide_(occupancySum, windowDays);
  const soldOutRatio = safeDivide_(soldOutDays, windowDays);
  const compressionRatio = safeDivide_(compressionDays, windowDays);
  const overbookingRatio = safeDivide_(overbookingDays, windowDays);
  const score = clampDemandScore_(
    (averageOccupancyPct * 65) +
    (soldOutRatio * 20) +
    (compressionRatio * 10) +
    (overbookingRatio * 5)
  );

  return {
    days: windowDays,
    startDate: formatDateKey_(targetDate),
    endDate: formatDateKey_(addDays_(targetDate, windowDays - 1)),
    score: score,
    band: getDemandBandFromScore_(score),
    averageOccupancyPct: averageOccupancyPct,
    soldOutDays: soldOutDays,
    compressionDays: compressionDays,
    overbookingDays: overbookingDays
  };
}

function calculateDemandScore_(exactDate, occupancyPct, competitorSummary, pickupMetrics, pressure7, eventFlagSummary) {
  const occupancyComponent = normalizeDemandOccupancyComponent_(occupancyPct, exactDate.availableRooms);
  const shortTermPressureComponent = Number(pressure7.score || 0) / 100;
  const competitorSoldOutComponent = Number(competitorSummary.totalCompetitorsChecked || 0) > 0
    ? safeDivide_(Number(competitorSummary.competitorSoldOutCount || 0), Number(competitorSummary.totalCompetitorsChecked || 0))
    : null;
  const marketGapComponent = normalizeDemandMarketGapComponent_(
    competitorSummary.rozaCurrentPrice,
    competitorSummary.averageCompetitorPrice
  );
  const pickupSignalRooms = pickupMetrics && pickupMetrics.pickupSignalRooms != null
    ? Number(pickupMetrics.pickupSignalRooms)
    : null;
  const pickupComponent = pickupSignalRooms != null
    ? normalizeDemandPickupComponent_(pickupSignalRooms, exactDate.sellableInventory)
    : null;
  const eventFlagComponent = eventFlagSummary && eventFlagSummary.hasActiveFlag
    ? getEventFlagImpactDemandComponent_(eventFlagSummary.highestImpactLevel)
    : null;

  const components = [
    { key: 'occupancy', weight: 40, value: occupancyComponent },
    { key: 'short_term_pressure', weight: 20, value: shortTermPressureComponent },
    { key: 'competitor_sold_out', weight: 15, value: competitorSoldOutComponent },
    { key: 'market_gap', weight: 15, value: marketGapComponent },
    { key: 'pickup_history', weight: 10, value: pickupComponent },
    { key: 'event_flag', weight: 10, value: eventFlagComponent }
  ];

  let totalWeight = 0;
  let weightedValue = 0;
  const usedSignals = [];
  const componentScores = {};

  components.forEach(function(component) {
    if (component.value == null) {
      componentScores[component.key] = null;
      return;
    }
    totalWeight += Number(component.weight || 0);
    weightedValue += Number(component.value || 0) * Number(component.weight || 0);
    usedSignals.push(component.key);
    componentScores[component.key] = clampDemandScore_(Number(component.value || 0) * 100);
  });

  return {
    score: totalWeight ? clampDemandScore_((weightedValue / totalWeight) * 100) : 0,
    usedSignals: usedSignals,
    componentScores: componentScores
  };
}

function normalizeDemandOccupancyComponent_(occupancyPct, availableRooms) {
  if (Number(availableRooms || 0) <= 0) return 1;
  return clampDemandRatio_(safeDivide_(Number(occupancyPct || 0), 0.9));
}

function normalizeDemandMarketGapComponent_(rozaCurrentPrice, competitorAveragePrice) {
  if (rozaCurrentPrice == null || competitorAveragePrice == null || Number(competitorAveragePrice || 0) <= 0) {
    return null;
  }
  const ratio = safeDivide_(Number(competitorAveragePrice || 0) - Number(rozaCurrentPrice || 0), Number(competitorAveragePrice || 0));
  return clampDemandRatio_(0.5 + ratio);
}

function normalizeDemandPickupComponent_(roomsDelta, sellableInventory) {
  const base = Math.max(1, Number(sellableInventory || 0));
  const ratio = safeDivide_(Number(roomsDelta || 0), base);
  return clampDemandRatio_(0.5 + ratio);
}

function hasPickupPaceBookingHistoryForStay_(bookingNightRows, targetDate, roomTypeId) {
  const stayKey = formatDateKey_(targetDate);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  return (bookingNightRows || []).some(function(row) {
    return String(row.stay_date || '').trim() === stayKey &&
      String(row.room_type_id || '').trim() === targetRoomTypeId &&
      !!normalizeDateInput_(row.booking_created_at);
  });
}

function getPickupPaceBookingWindowRooms_(bookingNightRows, targetDate, roomTypeId, asOfDate, daysBack) {
  const stayKey = formatDateKey_(targetDate);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const endDate = stripTime_(asOfDate || new Date());
  const windowDays = Math.max(1, Number(daysBack || 1));
  const startDate = addDays_(endDate, -(windowDays - 1));
  let rooms = 0;

  (bookingNightRows || []).forEach(function(row) {
    if (String(row.stay_date || '').trim() !== stayKey) return;
    if (String(row.room_type_id || '').trim() !== targetRoomTypeId) return;
    const createdAt = normalizeDateInput_(row.booking_created_at);
    if (!createdAt) return;
    if (createdAt.getTime() < startDate.getTime() || createdAt.getTime() > endDate.getTime()) return;
    rooms += Number(row.qty_rooms || row.room_nights || 0);
  });

  return Math.max(0, Number(rooms || 0));
}

function getPickupPaceBand_(signalRooms, sellableInventory) {
  const rooms = Number(signalRooms || 0);
  const inventory = Math.max(1, Number(sellableInventory || 0));
  if (rooms >= Math.max(2, Math.ceil(inventory * 0.5))) return 'Strong';
  if (rooms >= Math.max(1, Math.ceil(inventory * 0.25))) return 'Normal';
  return 'Weak';
}

function getPickupPaceTrend_(recentRooms3Days, priorRooms11Days) {
  const recent = Number(recentRooms3Days || 0);
  const prior = Number(priorRooms11Days || 0);
  if (recent <= 0 && prior <= 0) return 'Flat';
  if (prior <= 0 && recent > 0) return 'Improving';
  if (recent <= 0 && prior > 0) return 'Slowing';

  const recentRate = safeDivide_(recent, 3);
  const priorRate = safeDivide_(prior, 11);
  if (recentRate >= priorRate + 0.25) return 'Improving';
  if (recentRate <= Math.max(0, priorRate - 0.25)) return 'Slowing';
  return 'Flat';
}

function getPickupPaceSnapshotTrend_(paceDelta7Days, paceDelta14Days) {
  if (paceDelta7Days == null && paceDelta14Days == null) return '';
  if (paceDelta7Days == null) return '';
  if (paceDelta14Days == null) return paceDelta7Days > 0 ? 'Improving' : (paceDelta7Days < 0 ? 'Slowing' : 'Flat');
  const priorSevenDays = Number(paceDelta14Days || 0) - Number(paceDelta7Days || 0);
  const recentRate = safeDivide_(Number(paceDelta7Days || 0), 7);
  const priorRate = safeDivide_(priorSevenDays, 7);
  if (recentRate >= priorRate + 0.15) return 'Improving';
  if (recentRate <= priorRate - 0.15) return 'Slowing';
  return 'Flat';
}

function buildPickupPaceSnapshotNote_(snapshotDelta7Days, snapshotDelta14Days) {
  if (snapshotDelta7Days && snapshotDelta7Days.available) {
    return snapshotDelta7Days.note || 'Pace delta compares current live OTB with the closest older snapshot.';
  }
  if (snapshotDelta14Days && snapshotDelta14Days.available) {
    return snapshotDelta14Days.note || 'Pace delta compares current live OTB with the closest older snapshot.';
  }
  if (snapshotDelta7Days && snapshotDelta7Days.note) return snapshotDelta7Days.note;
  if (snapshotDelta14Days && snapshotDelta14Days.note) return snapshotDelta14Days.note;
  return 'Snapshot pace history is not available yet.';
}

function buildPickupPaceSnapshotSignal_(snapshotRows, asOfDate, targetDate, roomTypeId, daysBack) {
  const currentReferenceDate = getLatestSnapshotDateOnOrBefore_(snapshotRows, stripTime_(asOfDate || new Date()));
  if (!currentReferenceDate) {
    return {
      available: false,
      daysBack: Number(daysBack || 7),
      note: 'No OTB snapshot history is available yet.'
    };
  }

  const previousSnapshotDate = getLatestSnapshotDateOnOrBefore_(
    snapshotRows,
    addDays_(currentReferenceDate, -Math.max(1, Number(daysBack || 7)))
  );
  if (!previousSnapshotDate) {
    return {
      available: false,
      daysBack: Number(daysBack || 7),
      currentSnapshotDate: formatDateKey_(currentReferenceDate),
      note: 'An older OTB snapshot is needed before pace can be measured.'
    };
  }

  const currentMetrics = getDemandSnapshotMetric_(snapshotRows, currentReferenceDate, targetDate, roomTypeId);
  const previousMetrics = getDemandSnapshotMetric_(snapshotRows, previousSnapshotDate, targetDate, roomTypeId);
  if (!currentMetrics.found || !previousMetrics.found) {
    return {
      available: false,
      daysBack: Number(daysBack || 7),
      currentSnapshotDate: formatDateKey_(currentReferenceDate),
      previousSnapshotDate: formatDateKey_(previousSnapshotDate),
      note: 'Snapshot history does not cover this stay date yet.'
    };
  }

  return {
    available: true,
    daysBack: Number(daysBack || 7),
    currentSnapshotDate: formatDateKey_(currentReferenceDate),
    previousSnapshotDate: formatDateKey_(previousSnapshotDate),
    currentRoomsSold: Number(currentMetrics.roomsSold || 0),
    previousRoomsSold: Number(previousMetrics.roomsSold || 0),
    roomsDelta: Number(currentMetrics.roomsSold || 0) - Number(previousMetrics.roomsSold || 0),
    occupancyDeltaPct: Number(currentMetrics.occupancyPct || 0) - Number(previousMetrics.occupancyPct || 0),
    note: 'Compared with the closest snapshot at least ' + String(Math.max(1, Number(daysBack || 7))) + ' day(s) earlier.'
  };
}

function buildDemandPickupSignal_(snapshotRows, targetDate, roomTypeId, daysBack) {
  return buildPickupPaceSnapshotSignal_(snapshotRows, stripTime_(new Date()), targetDate, roomTypeId, daysBack);
}

function getLatestSnapshotDateOnOrBefore_(snapshotRows, referenceDate) {
  const targetDate = stripTime_(referenceDate);
  let selected = null;

  (snapshotRows || []).forEach(function(row) {
    const snapshotDate = normalizeDateInput_(row.snapshot_date);
    if (!snapshotDate || snapshotDate.getTime() > targetDate.getTime()) return;
    if (!selected || snapshotDate.getTime() > selected.getTime()) {
      selected = snapshotDate;
    }
  });

  return selected;
}

function getDemandSnapshotMetric_(snapshotRows, snapshotDate, stayDate, roomTypeId) {
  const snapshotKey = formatDateKey_(snapshotDate);
  const stayKey = formatDateKey_(stayDate);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const match = (snapshotRows || []).find(function(row) {
    return String(row.snapshot_date || '').trim() === snapshotKey &&
      String(row.stay_date || '').trim() === stayKey &&
      String(row.room_type_id || '').trim() === targetRoomTypeId;
  });

  if (!match) {
    return { found: false };
  }

  return {
    found: true,
    roomsSold: Number(match.rooms_sold_otb || 0),
    roomsAvailableToSell: Number(match.rooms_available_to_sell || 0),
    occupancyPct: Number(match.occupancy_pct_otb || 0),
    roomRevenue: Number(match.room_revenue_otb || 0)
  };
}

function getDemandBandFromScore_(score) {
  const numeric = clampDemandScore_(score);
  if (numeric >= 75) return 'Hot';
  if (numeric >= 50) return 'Strong';
  if (numeric >= 25) return 'Normal';
  return 'Low';
}

function getDemandBandRank_(band) {
  const normalized = String(band || '').trim().toLowerCase();
  if (normalized === 'hot') return 4;
  if (normalized === 'strong') return 3;
  if (normalized === 'normal') return 2;
  return 1;
}

function getExactSpecialCommercialControlForDate_(roomTypeId, date, prefetchedRows) {
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const targetKey = formatDateKey_(date);
  let selected = null;

  (prefetchedRows || getCommercialControlRows_({ activeOnly: false, roomTypeId: targetRoomTypeId })).forEach(function(row) {
    if (!row || row.roomTypeId !== targetRoomTypeId) return;
    if (normalizeCommercialRuleType_(row.ruleType) !== 'special') return;
    if (!row.startDate || !row.endDate) return;
    if (formatDateKey_(row.startDate) !== targetKey || formatDateKey_(row.endDate) !== targetKey) return;
    if (!selected) {
      selected = row;
      return;
    }
    const activeDelta = Number(Boolean(selected.active)) - Number(Boolean(row.active));
    if (activeDelta !== 0) {
      if (row.active) selected = row;
      return;
    }
    if (compareCommercialControlPriority_(row, selected) < 0) {
      selected = row;
    }
  });

  return selected;
}

function clampDemandRatio_(value) {
  return Math.max(0, Math.min(1, Number(value || 0)));
}

function clampDemandScore_(value) {
  return Math.max(0, Math.min(100, Math.round(Number(value || 0))));
}

function normalizeBaseRateApplyMode_(value) {
  const normalized = String(value || '').trim().toLowerCase().replace(/[\s-]+/g, '_');
  if (normalized === 'selected' || normalized === 'selected_room_types') return 'selected';
  if (normalized === 'all' || normalized === 'all_room_types') return 'all';
  return 'single';
}

function normalizeRoomTypeIdList_(value) {
  const rawList = Array.isArray(value)
    ? value
    : String(value == null ? '' : value).split(/[,\|]/);
  const seen = {};
  const ids = [];
  rawList.forEach(function(item) {
    const raw = String(item || '').trim();
    if (!raw) return;
    const roomTypeId = resolveRoomTypeId_(raw);
    if (seen[roomTypeId]) return;
    seen[roomTypeId] = true;
    ids.push(roomTypeId);
  });
  return ids;
}

function validateBaseRateInput_(input) {
  const payload = input || {};
  const applyMode = normalizeBaseRateApplyMode_(payload.apply_to || payload.applyTo || payload.base_rate_apply_to || payload.baseRateApplyTo || 'single');
  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || payload.room_type || payload.roomType || '').trim();
  const baseRateRaw = String(payload.base_rate == null ? '' : payload.base_rate).trim();
  const extraGuestFeeRaw = String(payload.extra_guest_fee == null ? '' : payload.extra_guest_fee).trim();
  const directDiscount = normalizeDirectBookingDiscountFields_(payload);
  const activeRoomTypes = getActiveRoomTypeCatalog_();
  const activeRoomTypeIds = activeRoomTypes.map(function(row) {
    return row.roomTypeId;
  });
  const selectedRoomTypeIds = normalizeRoomTypeIdList_(
    payload.selected_room_type_ids != null ? payload.selected_room_type_ids :
      (payload.selectedRoomTypeIds != null ? payload.selectedRoomTypeIds :
        (payload.bulk_room_type_ids != null ? payload.bulk_room_type_ids : payload.bulkRoomTypeIds))
  );
  let targetRoomTypeIds = [];
  let roomTypeId = '';
  let roomTypeName = '';
  let baseRate = '';
  let extraGuestFee = '';

  if (applyMode === 'single') {
    if (!roomTypeRaw) throw new Error('Room type is required.');
    if (baseRateRaw === '') throw new Error('Base rate is required.');
    roomTypeId = resolveRoomTypeId_(roomTypeRaw);
    roomTypeName = getRoomTypeNameById_(roomTypeId);
    baseRate = roundCurrency_(baseRateRaw);
    if (!isFinite(Number(baseRate)) || Number(baseRate) < 0) {
      throw new Error('Base rate must be a valid amount.');
    }
    extraGuestFee = extraGuestFeeRaw === '' ? '' : roundCurrency_(extraGuestFeeRaw);
    if (extraGuestFeeRaw !== '' && (!isFinite(Number(extraGuestFee)) || Number(extraGuestFee) < 0)) {
      throw new Error('Extra guest fee must be zero or higher.');
    }
    validateDirectBookingDiscountFields_(directDiscount, baseRate);
    targetRoomTypeIds = [roomTypeId];
  } else if (applyMode === 'all') {
    targetRoomTypeIds = activeRoomTypeIds.slice();
  } else {
    targetRoomTypeIds = selectedRoomTypeIds.filter(function(roomTypeId) {
      return activeRoomTypeIds.indexOf(roomTypeId) !== -1;
    });
    if (!targetRoomTypeIds.length) {
      throw new Error('Select at least one active room type for bulk discount apply.');
    }
  }

  if (applyMode !== 'single') {
    if (normalizeDirectDiscountType_(directDiscount.directDiscountType) !== 'None' && Number(directDiscount.directDiscountValue || 0) <= 0) {
      throw new Error('Direct booking discount value must be greater than zero.');
    }
    if (directDiscount.publicReferencePrice !== '' && (!isFinite(Number(directDiscount.publicReferencePrice)) || Number(directDiscount.publicReferencePrice) <= 0)) {
      throw new Error('Public reference price must be a positive amount.');
    }
  }

  return {
    applyMode: applyMode,
    targetRoomTypeIds: targetRoomTypeIds,
    roomTypeId: roomTypeId,
    roomTypeName: roomTypeName,
    baseRate: baseRate,
    extraGuestFee: extraGuestFee,
    directDiscount: directDiscount,
    publicReferencePrice: directDiscount.publicReferencePrice,
    directDiscountType: directDiscount.directDiscountType,
    directDiscountValue: directDiscount.directDiscountValue,
    active: payload.active == null ? true : isYesLike_(payload.active),
    selectedDate: normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || new Date())
  };
}

function getBaseRatesBulkRowsFromInput_(input) {
  const payload = input || {};
  if (Array.isArray(payload)) return payload;
  const rows = payload.rows || payload.base_rates || payload.baseRates || payload.items;
  if (Array.isArray(rows)) return rows;
  if (typeof rows === 'string' && rows.trim()) {
    try {
      const parsed = JSON.parse(rows);
      if (Array.isArray(parsed)) return parsed;
    } catch (error) {
      throw new Error('Bulk base-rate rows could not be parsed.');
    }
  }
  return [];
}

function validateBaseRatesBulkInput_(input) {
  const rows = getBaseRatesBulkRowsFromInput_(input);
  const selectedRows = rows.filter(function(row) {
    return row && isYesLike_(row.selected);
  });
  if (!selectedRows.length) {
    throw new Error('Tick at least one room type before saving bulk base rates.');
  }

  const validatedRows = [];
  const skipped = [];
  const seen = {};
  selectedRows.forEach(function(row) {
    try {
      const roomTypeId = resolveRoomTypeId_(row.room_type_id || row.roomTypeId || row.room_type || row.roomType || '');
      if (seen[roomTypeId]) {
        skipped.push({
          roomTypeId: roomTypeId,
          roomTypeName: getRoomTypeNameById_(roomTypeId),
          reason: 'Room type appears more than once in the selected bulk rows.'
        });
        return;
      }
      seen[roomTypeId] = true;

      const roomTypeName = getRoomTypeNameById_(roomTypeId);
      const baseRateRaw = String(row.base_rate == null ? (row.baseRate == null ? '' : row.baseRate) : row.base_rate).trim();
      const extraGuestFeeRaw = String(row.extra_guest_fee == null ? (row.extraGuestFee == null ? '' : row.extraGuestFee) : row.extra_guest_fee).trim();
      if (baseRateRaw === '') {
        skipped.push({
          roomTypeId: roomTypeId,
          roomTypeName: roomTypeName,
          reason: 'Base rate is required.'
        });
        return;
      }
      const baseRate = roundCurrency_(baseRateRaw);
      if (!isFinite(Number(baseRate)) || Number(baseRate) < 0) {
        skipped.push({
          roomTypeId: roomTypeId,
          roomTypeName: roomTypeName,
          reason: 'Base rate must be zero or higher.'
        });
        return;
      }
      const extraGuestFee = extraGuestFeeRaw === '' ? '' : roundCurrency_(extraGuestFeeRaw);
      if (extraGuestFeeRaw !== '' && (!isFinite(Number(extraGuestFee)) || Number(extraGuestFee) < 0)) {
        skipped.push({
          roomTypeId: roomTypeId,
          roomTypeName: roomTypeName,
          reason: 'Extra guest fee must be zero or higher.'
        });
        return;
      }

      const directDiscount = normalizeDirectBookingDiscountFields_(row);
      validateDirectBookingDiscountFields_(directDiscount, baseRate);
      const activeRaw = row.active == null ? row.is_active : row.active;

      validatedRows.push({
        roomTypeId: roomTypeId,
        baseRate: baseRate,
        extraGuestFee: extraGuestFee,
        publicReferencePrice: directDiscount.publicReferencePrice,
        directDiscountType: directDiscount.directDiscountType,
        directDiscountValue: directDiscount.directDiscountValue,
        active: activeRaw == null || activeRaw === '' ? null : isYesLike_(activeRaw)
      });
    } catch (error) {
      skipped.push({
        roomTypeId: String(row && (row.room_type_id || row.roomTypeId || row.room_type || row.roomType) || '').trim(),
        reason: error.message || String(error)
      });
    }
  });

  return {
    rows: validatedRows,
    skipped: skipped
  };
}

function validateCommercialControlInput_(input, options) {
  const config = options || {};
  const payload = input || {};
  const roomTypeRaw = String(payload.room_type_id || payload.roomTypeId || payload.room_type || payload.roomType || '').trim();
  const startDate = normalizeDateInput_(payload.start_date || payload.startDate);
  const endDate = normalizeDateInput_(payload.end_date || payload.endDate);
  const overrideRaw = String(payload.override_price == null ? '' : payload.override_price).trim();
  const overbookingRaw = String(payload.overbooking_allowance == null ? '' : payload.overbooking_allowance).trim();
  const directDiscount = normalizeDirectBookingDiscountFields_(payload);
  const activeRaw = payload.active;

  if (!roomTypeRaw) throw new Error('Room type is required.');
  if (!startDate) throw new Error('Start date is required.');
  if (!endDate) throw new Error('End date is required.');
  if (endDate.getTime() < startDate.getTime()) throw new Error('End date must be on or after start date.');

  const roomTypeId = resolveRoomTypeId_(roomTypeRaw);
  const overridePrice = overrideRaw === '' ? '' : roundCurrency_(overrideRaw);
  const overbookingAllowance = overbookingRaw === '' ? 0 : Math.max(0, Number(overbookingRaw));
  const note = String(payload.note || payload.notes || '').trim();

  if (overrideRaw !== '' && (!isFinite(Number(overridePrice)) || Number(overridePrice) < 0)) {
    throw new Error('Override price must be a valid positive amount.');
  }
  if (overbookingRaw !== '' && (!isFinite(Number(overbookingAllowance)) || Number(overbookingAllowance) < 0)) {
    throw new Error('Overbooking allowance must be zero or higher.');
  }
  if (overbookingRaw !== '' && !Number.isInteger(Number(overbookingRaw))) {
    throw new Error('Overbooking allowance must be a whole-room number.');
  }
  validateDirectBookingDiscountFields_(directDiscount);
  const hasDirectDiscountConfig = hasDirectBookingDiscountConfig_(directDiscount);
  if (overrideRaw === '' && overbookingAllowance === 0 && !hasDirectDiscountConfig) {
    throw new Error('Enter a live price, an overbooking allowance, a direct-booking discount, or a combination.');
  }
  if ((overrideRaw !== '' || overbookingAllowance > 0 || hasDirectDiscountConfig) && !note) {
    throw new Error('Add an operator note before saving a live override or overbooking change.');
  }

  return {
    controlId: String(payload.control_id || payload.controlId || '').trim(),
    roomTypeId: roomTypeId,
    roomTypeName: String((config.roomTypeNameMap && config.roomTypeNameMap[roomTypeId]) || getRoomTypeNameById_(roomTypeId)).trim(),
    ruleType: normalizeCommercialRuleType_(payload.rule_type || payload.ruleType),
    startDate: startDate,
    endDate: endDate,
    overridePrice: overrideRaw === '' ? '' : overridePrice,
    publicReferencePrice: directDiscount.publicReferencePrice,
    directDiscountType: directDiscount.directDiscountType,
    directDiscountValue: directDiscount.directDiscountValue,
    overbookingAllowance: overbookingAllowance,
    active: activeRaw == null ? true : isYesLike_(activeRaw),
    note: note,
    selectedDate: normalizeFrontDeskDate_(payload.selected_date || payload.selectedDate || new Date())
  };
}

function calculateEstimatedPrice_(checkIn, checkOut, roomTypeId, guests) {
  const resolved = resolveStayCommercialQuote_(checkIn, checkOut, roomTypeId, guests);
  return {
    total: resolved.missingRateDates.length ? null : resolved.total,
    missingRateDates: resolved.missingRateDates,
    notes: resolved.notes,
    pricingSource: resolved.pricingSource,
    pricingReferenceId: resolved.pricingReferenceId,
    nightly: resolved.nightly
  };
}

function getOpenRateRowsForType_(roomTypeId) {
  return getSheetObjects_(SHEET_NAMES.RATES).filter(function(row) {
    return String(row.room_type_id || '').trim() === roomTypeId && String(row.status || '').trim().toLowerCase() === 'open';
  });
}

function getLegacyRateFallbackInfo_(roomTypeId, date, prefetchedRows) {
  const targetKey = formatDateKey_(date);
  const rows = prefetchedRows || [];

  if (!rows.length) {
    return { found: false };
  }

  let exact = null;
  let latestPast = null;
  let earliestFuture = null;

  rows.forEach(function(row) {
    const rateDate = normalizeDateInput_(row.date);
    if (!rateDate) return;
    const rateKey = formatDateKey_(rateDate);

    if (rateKey === targetKey) {
      exact = row;
      return;
    }

    if (rateDate.getTime() < stripTime_(date).getTime()) {
      if (!latestPast || rateDate.getTime() > normalizeDateInput_(latestPast.date).getTime()) {
        latestPast = row;
      }
      return;
    }

    if (!earliestFuture || rateDate.getTime() < normalizeDateInput_(earliestFuture.date).getTime()) {
      earliestFuture = row;
    }
  });

  const selected = exact || latestPast || earliestFuture;
  if (!selected) return { found: false };

  return {
    found: true,
    source: exact ? 'exact' : (latestPast ? 'latest-past' : 'earliest-future'),
    baseRate: Number(selected.base_rate || 0),
    extraGuestFee: Number(selected.extra_guest_fee || 0),
    minStay: Number(selected.min_stay || 1)
  };
}

function getBaseRateForType_(roomTypeId, prefetchedRows, legacyRateRows) {
  const rows = prefetchedRows || getBaseRateRows_({ activeOnly: true, roomTypeId: roomTypeId });
  let selected = null;
  rows.forEach(function(row) {
    if (!row.active || row.roomTypeId !== roomTypeId) return;
    if (!selected || row.updatedAtMs > selected.updatedAtMs || (row.updatedAtMs === selected.updatedAtMs && row.rowNumber > selected.rowNumber)) {
      selected = row;
    }
  });

  if (selected) {
    return {
      found: true,
      baseRate: Number(selected.baseRate || 0),
      extraGuestFee: selected.hasExtraGuestFee ? Number(selected.extraGuestFee || 0) : '',
      hasExtraGuestFee: selected.hasExtraGuestFee,
      publicReferencePrice: selected.publicReferencePrice,
      directDiscountType: selected.directDiscountType,
      directDiscountValue: selected.directDiscountValue,
      pricingSource: 'base_rate',
      pricingReferenceId: selected.roomTypeId
    };
  }

  const fallbackRate = getLegacyFallbackBaseRateForRoomType_(roomTypeId);
  if (fallbackRate > 0) {
    return {
      found: true,
      baseRate: fallbackRate,
      extraGuestFee: '',
      hasExtraGuestFee: false,
      publicReferencePrice: '',
      directDiscountType: 'None',
      directDiscountValue: '',
      pricingSource: 'base_rate',
      pricingReferenceId: roomTypeId
    };
  }

  const legacyFallback = getLegacyRateFallbackInfo_(roomTypeId, new Date(), legacyRateRows || getOpenRateRowsForType_(roomTypeId));
  if (legacyFallback.found) {
    return {
      found: true,
      baseRate: Number(legacyFallback.baseRate || 0),
      extraGuestFee: Number(legacyFallback.extraGuestFee || 0),
      hasExtraGuestFee: true,
      publicReferencePrice: '',
      directDiscountType: 'None',
      directDiscountValue: '',
      pricingSource: 'base_rate',
      pricingReferenceId: roomTypeId
    };
  }

  return { found: false };
}

function getRateCalendarInheritedDirectDiscountFields_(roomTypeId, overridePrice) {
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const baseRateInfo = getBaseRateForType_(
    targetRoomTypeId,
    getBaseRateRows_({ activeOnly: true, roomTypeId: targetRoomTypeId })
  );
  const inherited = normalizeDirectBookingDiscountFields_(baseRateInfo);
  if (!hasDirectBookingDiscountValue_(inherited)) {
    return {
      publicReferencePrice: '',
      directDiscountType: 'None',
      directDiscountValue: ''
    };
  }

  const effectiveReferencePrice = inherited.publicReferencePrice === ''
    ? normalizeOptionalPositiveCurrency_(overridePrice)
    : inherited.publicReferencePrice;

  return {
    publicReferencePrice: effectiveReferencePrice === '' ? '' : effectiveReferencePrice,
    directDiscountType: inherited.directDiscountType,
    directDiscountValue: inherited.directDiscountValue === '' ? '' : inherited.directDiscountValue
  };
}

function getDirectBookingDiscountSource_(priceControl, baseRateInfo) {
  if (priceControl && hasDirectBookingDiscountConfig_(priceControl)) return priceControl;
  if (!priceControl && baseRateInfo && hasDirectBookingDiscountConfig_(baseRateInfo, baseRateInfo.baseRate)) return baseRateInfo;
  return null;
}

function calculateDirectBookingDiscountedBaseRate_(discountSource, fallbackReferencePrice) {
  const fields = normalizeDirectBookingDiscountFields_(discountSource);
  if (!hasDirectBookingDiscountConfig_(fields, fallbackReferencePrice)) return null;
  const referencePrice = getDirectBookingReferencePrice_(fields, fallbackReferencePrice);
  const discountAmount = fields.directDiscountType === 'Percent'
    ? roundCurrency_(referencePrice * Number(fields.directDiscountValue || 0) / 100)
    : roundCurrency_(fields.directDiscountValue || 0);
  const directBaseRate = roundCurrency_(referencePrice - discountAmount);
  if (!(directBaseRate > 0) || !(referencePrice > directBaseRate)) return null;
  return {
    publicReferencePrice: referencePrice,
    directBaseRate: directBaseRate,
    discountAmount: discountAmount,
    directDiscountType: fields.directDiscountType,
    directDiscountValue: fields.directDiscountValue
  };
}

function buildDirectBookingNightOffer_(directNightlyRate, extraGuestFeeApplied, discountSource) {
  if (directNightlyRate == null) return null;
  const fields = normalizeDirectBookingDiscountFields_(discountSource);
  if (!hasDirectBookingDiscountConfig_(fields)) return null;
  const referenceNightlyRate = roundCurrency_(Number(fields.publicReferencePrice || 0) + Number(extraGuestFeeApplied || 0));
  const directRate = roundCurrency_(directNightlyRate);
  if (!(referenceNightlyRate > directRate)) return null;
  const savingsAmount = roundCurrency_(referenceNightlyRate - directRate);
  return {
    publicReferenceRate: referenceNightlyRate,
    directPrice: directRate,
    directDiscountType: fields.directDiscountType,
    directDiscountValue: fields.directDiscountValue,
    savingsAmount: savingsAmount,
    savingsPercentage: Math.round(safeDivide_(savingsAmount, referenceNightlyRate) * 10000) / 100
  };
}

function formatDirectBookingSavingsPercentLabel_(value) {
  const rounded = Math.round(Number(value || 0) * 10) / 10;
  return String(rounded).replace(/\.0$/, '') + '%';
}

function buildDirectBookingSavingsLabel_(offer, currency) {
  const details = offer || {};
  if (Number(details.savingsAmount || 0) <= 0) return '';
  if (String(details.directDiscountType || '') === 'Percent') {
    return 'Save ' + formatDirectBookingSavingsPercentLabel_(details.savingsPercentage) + ' when booking direct';
  }
  return 'Save ' + formatMoneyValue_(details.savingsAmount, currency || getDefaultBookingCurrency_()) + ' when booking direct';
}

function buildStayDirectBookingOffer_(nightlyRows, directTotal, qtyRooms, currency) {
  const nightly = nightlyRows || [];
  const roomQty = Math.max(1, Number(qtyRooms || 1));
  const directStayTotal = directTotal == null ? null : roundCurrency_(directTotal);
  if (directStayTotal == null || !nightly.length) {
    return {
      hasSaving: false,
      directPrice: directStayTotal
    };
  }
  const qualifyingRows = nightly.filter(function(row) {
    return row && Number(row.publicReferenceRate || 0) > Number(row.finalNightlyRate || 0);
  });
  if (qualifyingRows.length !== nightly.length) {
    return {
      hasSaving: false,
      directPrice: directStayTotal
    };
  }

  const referencePerRoom = qualifyingRows.reduce(function(sum, row) {
    return sum + Number(row.publicReferenceRate || 0);
  }, 0);
  const referenceTotal = roundCurrency_(referencePerRoom * roomQty);
  if (!(referenceTotal > directStayTotal)) {
    return {
      hasSaving: false,
      directPrice: directStayTotal
    };
  }
  const savingsAmount = roundCurrency_(referenceTotal - directStayTotal);
  const savingsPercentage = Math.round(safeDivide_(savingsAmount, referenceTotal) * 10000) / 100;
  const firstType = String(qualifyingRows[0].directDiscountType || 'None');
  const sameType = qualifyingRows.every(function(row) {
    return String(row.directDiscountType || 'None') === firstType;
  });
  const firstValue = qualifyingRows[0].directDiscountValue;
  const sameValue = qualifyingRows.every(function(row) {
    return String(row.directDiscountValue || '') === String(firstValue || '');
  });
  const offer = {
    hasSaving: true,
    directPrice: directStayTotal,
    comparisonPrice: referenceTotal,
    publicReferencePrice: referenceTotal,
    savingsAmount: savingsAmount,
    savingsPercentage: savingsPercentage,
    directDiscountType: sameType ? firstType : 'Mixed',
    directDiscountValue: sameType && sameValue ? firstValue : '',
    currency: currency || getDefaultBookingCurrency_()
  };
  offer.savingsLabel = buildDirectBookingSavingsLabel_(offer, offer.currency);
  return offer;
}

function resolveLiveCommercialDate_(date, roomTypeId, guests, options) {
  const config = options || {};
  const normalizedDate = normalizeDateInput_(date);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const physicalInventory = Number(config.physicalInventory != null ? config.physicalInventory : getRoomInventory_(targetRoomTypeId));
  const commercialControls = config.commercialControls || getCommercialControlRows_({ activeOnly: true, roomTypeId: targetRoomTypeId });
  const baseRateRows = config.baseRateRows || getBaseRateRows_({ activeOnly: true, roomTypeId: targetRoomTypeId });
  const legacyRateRows = config.legacyRateRows || getOpenRateRowsForType_(targetRoomTypeId);
  const bookingCountMap = config.bookingCountMap;
  const blockedCountMap = config.blockedCountMap;

  const baseRateInfo = getBaseRateForType_(targetRoomTypeId, baseRateRows, legacyRateRows);
  const priceControl = getCommercialPriceControlForDate_(targetRoomTypeId, normalizedDate, commercialControls);
  const overbookingControl = getCommercialOverbookingControlForDate_(targetRoomTypeId, normalizedDate, commercialControls);
  const legacyRateInfo = getLegacyRateFallbackInfo_(targetRoomTypeId, normalizedDate, legacyRateRows);
  const includedGuests = Number(INCLUDED_GUESTS_BASELINE_BY_ROOM_TYPE[targetRoomTypeId] || 2);
  const extraGuests = Math.max(0, Number(guests || 1) - includedGuests);
  const extraGuestFee = baseRateInfo.hasExtraGuestFee
    ? Number(baseRateInfo.extraGuestFee || 0)
    : Number(legacyRateInfo.extraGuestFee || 0);
  const overrideRateUsed = priceControl && priceControl.hasOverridePrice ? Number(priceControl.overridePrice || 0) : null;
  const baseRate = Number(baseRateInfo.baseRate || 0);
  const nightlyBase = overrideRateUsed != null ? overrideRateUsed : baseRate;
  const directBookingDiscountSource = getDirectBookingDiscountSource_(priceControl, baseRateInfo);
  const directBookingDiscount = calculateDirectBookingDiscountedBaseRate_(
    directBookingDiscountSource,
    directBookingDiscountSource === baseRateInfo ? baseRate : null
  );
  const directNightlyBase = directBookingDiscount ? directBookingDiscount.directBaseRate : nightlyBase;
  const extraGuestFeeApplied = extraGuests > 0 && extraGuestFee > 0 ? extraGuests * extraGuestFee : 0;
  const finalNightlyRate = directNightlyBase + extraGuestFeeApplied;
  const finalNightlyRateValue = baseRateInfo.found || overrideRateUsed != null ? roundCurrency_(finalNightlyRate) : null;
  const directBookingNightOffer = buildDirectBookingNightOffer_(
    finalNightlyRateValue,
    extraGuestFeeApplied,
    directBookingDiscount
  );
  const soldRooms = getConfirmedBookingsCount_(normalizedDate, targetRoomTypeId, bookingCountMap);
  const blockedRooms = getBlockedDatesCount_(normalizedDate, targetRoomTypeId, blockedCountMap);
  const overbookingAllowanceApplied = Number(overbookingControl && overbookingControl.overbookingAllowance || 0);
  const sellableInventory = physicalInventory + overbookingAllowanceApplied;
  const availableRooms = sellableInventory - soldRooms - blockedRooms;
  const pricingSource = priceControl
    ? (priceControl.ruleType === 'special' ? 'special_override' : 'seasonal_override')
    : 'base_rate';
  const pricingReferenceId = priceControl ? priceControl.controlId : (baseRateInfo.pricingReferenceId || targetRoomTypeId);

  return {
    roomTypeId: targetRoomTypeId,
    date: formatDateKey_(normalizedDate),
    baseRate: baseRateInfo.found ? baseRate : null,
    overrideRateUsed: overrideRateUsed,
    directBaseRate: directBookingDiscount ? directBookingDiscount.directBaseRate : null,
    finalNightlyRate: finalNightlyRateValue,
    publicReferenceRate: directBookingNightOffer ? directBookingNightOffer.publicReferenceRate : null,
    directDiscountType: directBookingNightOffer ? directBookingNightOffer.directDiscountType : 'None',
    directDiscountValue: directBookingNightOffer ? directBookingNightOffer.directDiscountValue : '',
    savingsAmount: directBookingNightOffer ? directBookingNightOffer.savingsAmount : 0,
    savingsPercentage: directBookingNightOffer ? directBookingNightOffer.savingsPercentage : 0,
    pricingSource: pricingSource,
    pricingReferenceId: pricingReferenceId,
    physicalInventory: physicalInventory,
    overbookingAllowanceApplied: overbookingAllowanceApplied,
    sellableInventory: sellableInventory,
    soldRooms: soldRooms,
    blockedRooms: blockedRooms,
    availableRooms: availableRooms,
    extraGuestFeeApplied: extraGuestFeeApplied
  };
}

function resolveStayCommercialQuote_(checkIn, checkOut, roomTypeId, guests, options) {
  const stayDates = enumerateStayDates_(checkIn, checkOut);
  const config = options || {};
  const commercialControls = config.commercialControls || getCommercialControlRows_({ activeOnly: true, roomTypeId: roomTypeId });
  const baseRateRows = config.baseRateRows || getBaseRateRows_({ activeOnly: true, roomTypeId: roomTypeId });
  const legacyRateRows = config.legacyRateRows || getOpenRateRowsForType_(roomTypeId);
  const physicalInventory = Number(config.physicalInventory != null ? config.physicalInventory : getRoomInventory_(roomTypeId));
  const bookingRows = config.bookingRows || null;
  const blockedRows = config.blockedRows || null;
  const mapBounds = {
    startDate: checkIn,
    endDateExclusive: checkOut
  };
  const bookingCountMap = config.bookingCountMap || buildConfirmedBookingRoomCountMap_(bookingRows || getSheetObjects_(SHEET_NAMES.BOOKINGS), mapBounds).byDateRoom;
  const blockedCountMap = config.blockedCountMap || buildBlockedDateQtyMap_(blockedRows || getSheetObjects_(SHEET_NAMES.BLOCKED_DATES), mapBounds).byDateRoom;
  const missingRateDates = [];
  const notes = [];
  const nightly = [];
  let total = 0;

  stayDates.forEach(function(date) {
    const resolved = resolveLiveCommercialDate_(date, roomTypeId, guests, {
      commercialControls: commercialControls,
      baseRateRows: baseRateRows,
      legacyRateRows: legacyRateRows,
      physicalInventory: physicalInventory,
      bookingCountMap: bookingCountMap,
      blockedCountMap: blockedCountMap
    });

    if (resolved.finalNightlyRate == null) {
      missingRateDates.push(formatDateKey_(date));
      return;
    }

    total += Number(resolved.finalNightlyRate || 0);
    nightly.push(resolved);
    if (resolved.pricingSource !== 'base_rate') {
      notes.push('Used ' + resolved.pricingSource.replace(/_/g, ' ') + ' for ' + formatDateKey_(date));
    }
  });

  const sourceSet = nightly.reduce(function(map, row) {
    map[row.pricingSource] = true;
    return map;
  }, {});
  const referenceSet = nightly.reduce(function(map, row) {
    map[String(row.pricingReferenceId || '')] = true;
    return map;
  }, {});

  return {
    total: missingRateDates.length ? null : total,
    missingRateDates: missingRateDates,
    notes: notes,
    nightly: nightly,
    pricingSource: Object.keys(sourceSet).length <= 1 ? (nightly[0] ? nightly[0].pricingSource : 'base_rate') : 'mixed',
    pricingReferenceId: Object.keys(referenceSet).length <= 1 ? (nightly[0] ? nightly[0].pricingReferenceId : '') : 'mixed'
  };
}

function buildStayAvailabilityInventorySnapshot_(checkIn, checkOut, roomTypeId, guests, options) {
  const config = options || {};
  const resolvedRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const qtyRooms = Math.max(1, Number(config.qtyRooms || config.qty_rooms || 1));
  const guestsPerRoom = getGuestsPerRoom_(guests, qtyRooms);
  const requestedBedSetup = normalizeBedSetup_(config.bedSetup || config.bed_setup || '');
  const roomIndex = config.roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const sellableProductId = String(config.sellableProductId || config.sellable_product_id || '').trim();
  const sellableProduct = sellableProductId ? getSellableRoomProductById_(sellableProductId) : null;
  const sellableConstraint = getSellableProductConstraint_(sellableProduct, requestedBedSetup, roomIndex);
  const effectiveRequestedBedSetup = sellableConstraint.appliedBedSetup || requestedBedSetup;
  const requestedCandidateRoomIds = Array.isArray(config.candidateRoomIds) && config.candidateRoomIds.length
    ? config.candidateRoomIds.slice()
    : sellableConstraint.candidateRoomIds.slice();
  const breakdown = buildAvailabilityBreakdown_(checkIn, checkOut, resolvedRoomTypeId, {
    guests: guestsPerRoom,
    bedSetup: effectiveRequestedBedSetup,
    candidateRoomIds: requestedCandidateRoomIds,
    excludeBookingId: String(config.excludeBookingId || config.exclude_booking_id || '').trim(),
    inventoryTotal: config.inventoryTotal != null ? config.inventoryTotal : config.physicalInventory,
    bookingRows: config.bookingRows,
    blockedRows: config.blockedRows,
    bookingCountMap: config.bookingCountMap,
    blockedCountMap: config.blockedCountMap,
    demandUnitsByDateRoom: config.demandUnitsByDateRoom,
    roomIndex: roomIndex
  });
  const inventoryFallback = breakdown.length
    ? Number(breakdown[0].totalRooms || 0)
    : Number(config.inventoryTotal != null ? config.inventoryTotal : (config.physicalInventory != null ? config.physicalInventory : getRoomInventory_(resolvedRoomTypeId)));
  const totalRooms = breakdown.reduce(function(min, row) {
    return Math.min(min, Number(row.totalRooms || 0));
  }, inventoryFallback);
  const soldRooms = breakdown.reduce(function(max, row) {
    return Math.max(max, Number(row.soldRooms || 0));
  }, 0);
  const blockedRooms = breakdown.reduce(function(max, row) {
    return Math.max(max, Number(row.blockedRooms || 0));
  }, 0);
  const availableRooms = breakdown.reduce(function(min, row) {
    return Math.min(min, Number(row.availableRooms || 0));
  }, totalRooms);

  return {
    roomTypeId: resolvedRoomTypeId,
    roomTypeName: String((config.roomTypeNameMap && config.roomTypeNameMap[resolvedRoomTypeId]) || getRoomTypeNameById_(resolvedRoomTypeId)).trim(),
    qtyRooms: qtyRooms,
    guestsPerRoom: guestsPerRoom,
    breakdown: breakdown,
    bedSetup: effectiveRequestedBedSetup || 'Best available',
    source: breakdown.length && breakdown[0].source ? breakdown[0].source : 'live',
    totalRooms: totalRooms,
    soldRooms: soldRooms,
    blockedRooms: blockedRooms,
    availableRooms: availableRooms,
    hasInventoryConfigured: totalRooms > 0,
    roomIndex: roomIndex
  };
}

function buildStayAvailabilityPricingSnapshot_(checkIn, checkOut, roomTypeId, guests, options) {
  const config = options || {};
  const timingAccumulator = config.timingAccumulator || null;
  const inventoryStartedAt = Date.now();
  const inventory = buildStayAvailabilityInventorySnapshot_(checkIn, checkOut, roomTypeId, guests, config);
  if (timingAccumulator) {
    timingAccumulator.inventoryMs = Number(timingAccumulator.inventoryMs || 0) + (Date.now() - inventoryStartedAt);
  }
  const pricingStartedAt = Date.now();
  const pricing = resolveStayCommercialQuote_(checkIn, checkOut, inventory.roomTypeId, inventory.guestsPerRoom, {
    bookingRows: config.bookingRows,
    blockedRows: config.blockedRows,
    bookingCountMap: config.bookingCountMap,
    blockedCountMap: config.blockedCountMap,
    commercialControls: config.commercialControls,
    baseRateRows: config.baseRateRows,
    legacyRateRows: config.legacyRateRows,
    physicalInventory: config.physicalInventory != null ? config.physicalInventory : config.inventoryTotal
  });
  const pricingByDate = pricing.nightly.reduce(function(map, row) {
    map[row.date] = row;
    return map;
  }, {});
  const estimatedPricePerRoom = pricing.total == null ? null : roundCurrency_(pricing.total);
  const estimatedPrice = estimatedPricePerRoom == null ? null : roundCurrency_(estimatedPricePerRoom * inventory.qtyRooms);
  const directBookingOffer = buildStayDirectBookingOffer_(
    pricing.nightly,
    estimatedPrice,
    inventory.qtyRooms,
    config.currency || getDefaultBookingCurrency_()
  );
  if (timingAccumulator) {
    timingAccumulator.pricingMs = Number(timingAccumulator.pricingMs || 0) + (Date.now() - pricingStartedAt);
  }

  return {
    roomTypeId: inventory.roomTypeId,
    roomTypeName: inventory.roomTypeName,
    qtyRooms: inventory.qtyRooms,
    guestsPerRoom: inventory.guestsPerRoom,
    breakdown: inventory.breakdown,
    pricingByDate: pricingByDate,
    bedSetup: inventory.bedSetup,
    source: inventory.source,
    totalRooms: inventory.totalRooms,
    soldRooms: inventory.soldRooms,
    blockedRooms: inventory.blockedRooms,
    availableRooms: inventory.availableRooms,
    hasInventoryConfigured: inventory.hasInventoryConfigured,
    estimatedPrice: estimatedPrice,
    estimatedPricePerRoom: estimatedPricePerRoom,
    directBookingOffer: directBookingOffer,
    pricingSource: pricing.pricingSource,
    pricingReferenceId: pricing.pricingReferenceId,
    pricingNotes: pricing.notes,
    missingRateDates: pricing.missingRateDates
  };
}

function getRoomInventory_(roomTypeId) {
  const inventoryMap = getRoomInventoryMap_();
  if (Object.prototype.hasOwnProperty.call(inventoryMap, roomTypeId)) {
    return Number(inventoryMap[roomTypeId] || 0);
  }
  const row = getRoomTypeRowById_(roomTypeId);
  return Number(row.inventory_total || 0);
}

function getRoomTypeRowById_(roomTypeId) {
  const rows = getSheetObjects_(SHEET_NAMES.ROOM_TYPES);
  const match = rows.find(function(row) {
    return String(row.room_type_id || '').trim() === roomTypeId;
  });

  if (!match) {
    throw new Error('Room type not found: ' + roomTypeId);
  }

  return match;
}

function getRoomTypeNameById_(roomTypeId) {
  return String(getRoomTypeRowById_(roomTypeId).room_type_name || roomTypeId).trim();
}

function getConfirmedBookingsCount_(date, roomTypeId, bookingCountMap) {
  const key = getDateRoomKey_(date, roomTypeId);
  if (bookingCountMap) {
    return Number(bookingCountMap[key] || 0);
  }

  const targetDate = stripTime_(date).getTime();
  const rows = getSheetObjects_(SHEET_NAMES.BOOKINGS);
  let count = 0;

  rows.forEach(function(row) {
    const status = normalizeBookingStatus_(row.status);
    const currentRoomTypeId = String(row.room_type_id || '').trim();
    if (BOOKING_STATUSES_COUNTED.indexOf(status) === -1 || currentRoomTypeId !== roomTypeId) return;

    const checkIn = normalizeDateInput_(row.check_in);
    const checkOut = getInventoryEffectiveCheckOut_(row);
    if (!checkIn || !checkOut) return;

    const checkInMs = stripTime_(checkIn).getTime();
    const checkOutMs = stripTime_(checkOut).getTime();

    if (targetDate >= checkInMs && targetDate < checkOutMs) {
      count += Math.max(1, Number(row.qty_rooms || 1));
    }
  });

  return count;
}

function getBlockedDatesCount_(date, roomTypeId, blockedCountMap) {
  const key = getDateRoomKey_(date, roomTypeId);
  if (blockedCountMap) {
    return Number(blockedCountMap[key] || 0);
  }

  const targetDate = stripTime_(date).getTime();
  const rows = getSheetObjects_(SHEET_NAMES.BLOCKED_DATES);
  let blocked = 0;

  rows.forEach(function(row) {
    const currentRoomTypeId = String(row.room_type_id || '').trim();
    if (currentRoomTypeId !== roomTypeId) return;

    const status = String(row.status || 'Active').trim().toLowerCase();
    if (status && status !== 'active' && status !== 'open') return;

    const blockDate = normalizeDateInput_(row.date);
    if (!blockDate) return;

    if (stripTime_(blockDate).getTime() === targetDate) {
      blocked += Number(row.qty_blocked || 0);
    }
  });

  return blocked;
}

function getDateRoomKey_(date, roomTypeId) {
  return formatDateKey_(date) + '|' + String(roomTypeId || '').trim();
}

function buildConfirmedBookingRoomCountMap_(bookingRows, options) {
  const config = options || {};
  const rangeStart = normalizeDateInput_(config.startDate || config.start_date || config.fromDate || config.from_date);
  const rangeEndExclusive = normalizeDateInput_(config.endDateExclusive || config.end_date_exclusive || config.endDate || config.end_date || config.toDate || config.to_date);
  const byDateRoom = {};
  let minDate = null;
  let maxDate = null;

  (bookingRows || []).forEach(function(row) {
    const status = normalizeBookingStatus_(row.status);
    if (!isRevenueCountedStatus_(status)) return;

    const roomTypeId = String(row.room_type_id || '').trim();
    if (!roomTypeId) return;

    const checkIn = normalizeDateInput_(row.check_in);
    const checkOut = getInventoryEffectiveCheckOut_(row);
    if (!checkIn || !checkOut || checkOut.getTime() <= checkIn.getTime()) return;
    if (rangeStart && checkOut.getTime() <= rangeStart.getTime()) return;
    if (rangeEndExclusive && checkIn.getTime() >= rangeEndExclusive.getTime()) return;

    const qtyRooms = Math.max(1, Number(row.qty_rooms || 1));
    const effectiveCheckIn = rangeStart && checkIn.getTime() < rangeStart.getTime() ? rangeStart : checkIn;
    const effectiveCheckOut = rangeEndExclusive && checkOut.getTime() > rangeEndExclusive.getTime() ? rangeEndExclusive : checkOut;

    enumerateStayDates_(effectiveCheckIn, effectiveCheckOut).forEach(function(stayDate) {
      const key = getDateRoomKey_(stayDate, roomTypeId);
      byDateRoom[key] = Number(byDateRoom[key] || 0) + qtyRooms;
      if (!minDate || stayDate.getTime() < minDate.getTime()) minDate = stayDate;
      if (!maxDate || stayDate.getTime() > maxDate.getTime()) maxDate = stayDate;
    });
  });

  return {
    byDateRoom: byDateRoom,
    minDate: minDate,
    maxDate: maxDate
  };
}

function buildBlockedDateQtyMap_(blockedRows, options) {
  const config = options || {};
  const rangeStart = normalizeDateInput_(config.startDate || config.start_date || config.fromDate || config.from_date);
  const rangeEndExclusive = normalizeDateInput_(config.endDateExclusive || config.end_date_exclusive || config.endDate || config.end_date || config.toDate || config.to_date);
  const byDateRoom = {};
  let minDate = null;
  let maxDate = null;

  (blockedRows || []).forEach(function(row) {
    const roomTypeId = String(row.room_type_id || '').trim();
    if (!roomTypeId) return;

    const status = String(row.status || 'Active').trim().toLowerCase();
    if (status && status !== 'active' && status !== 'open') return;

    const blockDate = normalizeDateInput_(row.date);
    if (!blockDate) return;
    if (rangeStart && blockDate.getTime() < rangeStart.getTime()) return;
    if (rangeEndExclusive && blockDate.getTime() >= rangeEndExclusive.getTime()) return;

    const key = getDateRoomKey_(blockDate, roomTypeId);
    byDateRoom[key] = Number(byDateRoom[key] || 0) + Number(row.qty_blocked || 0);

    if (!minDate || blockDate.getTime() < minDate.getTime()) minDate = blockDate;
    if (!maxDate || blockDate.getTime() > maxDate.getTime()) maxDate = blockDate;
  });

  return {
    byDateRoom: byDateRoom,
    minDate: minDate,
    maxDate: maxDate
  };
}

function buildBookingNightStatsMap_(bookingNightRows) {
  const byDateRoom = {};
  let minDate = null;
  let maxDate = null;

  (bookingNightRows || []).forEach(function(row) {
    const stayDate = normalizeDateInput_(row.stay_date);
    if (!stayDate) return;

    const roomTypeId = String(row.room_type_id || '').trim();
    if (!roomTypeId) return;

    const key = getDateRoomKey_(stayDate, roomTypeId);
    if (!byDateRoom[key]) {
      byDateRoom[key] = {
        roomsSold: 0,
        roomNights: 0,
        roomRevenue: 0,
        bookingIds: {}
      };
    }

    byDateRoom[key].roomsSold += Number(row.room_nights || 0);
    byDateRoom[key].roomNights += Number(row.room_nights || 0);
    byDateRoom[key].roomRevenue += Number(row.nightly_room_revenue || 0);
    byDateRoom[key].bookingIds[String(row.booking_id || '')] = true;

    if (!minDate || stayDate.getTime() < minDate.getTime()) minDate = stayDate;
    if (!maxDate || stayDate.getTime() > maxDate.getTime()) maxDate = stayDate;
  });

  Object.keys(byDateRoom).forEach(function(key) {
    byDateRoom[key].bookingsCount = Object.keys(byDateRoom[key].bookingIds).filter(Boolean).length;
    delete byDateRoom[key].bookingIds;
  });

  return {
    byDateRoom: byDateRoom,
    minDate: minDate,
    maxDate: maxDate
  };
}

function enumerateStayDates_(checkIn, checkOut) {
  const start = stripTime_(normalizeDateInput_(checkIn));
  const end = stripTime_(normalizeDateInput_(checkOut));
  const dates = [];
  const cursor = new Date(start);

  while (cursor.getTime() < end.getTime()) {
    dates.push(new Date(cursor));
    cursor.setDate(cursor.getDate() + 1);
  }

  return dates;
}

function buildAvailabilityMessage_(available, missingRateDates, remainingMin) {
  if (missingRateDates.length) {
    return 'Availability could not be priced because some dates are missing rates.';
  }
  if (!available) {
    return remainingMin <= 0 ? 'Selected room type is sold out for at least one night in this stay.' : 'Requested stay is not available.';
  }
  return 'Requested stay is available.';
}

function buildWhatsappMessage_(details) {
  const checkIn = formatDateForGuest_(details.checkIn);
  const checkOut = formatDateForGuest_(details.checkOut);
  return 'Hi Roza! I\'d like to check availability for the ' + details.roomTypeName +
    ' from ' + checkIn + ' to ' + checkOut + ' for ' + details.guests +
    ' people. Preferred bed setup: ' + (details.bedSetup || 'Best available') + '. Is it available?';
}

function buildRequestAvailabilitySnapshot_(validated, roomTypeId, roomTypeName) {
  try {
    return checkAvailability({
      checkIn: validated.checkIn,
      checkOut: validated.checkOut,
      roomType: validated.roomType,
      guests: validated.guests,
      bedSetup: validated.bedSetup
    });
  } catch (error) {
    return {
      available: null,
      nights: enumerateStayDates_(validated.checkIn, validated.checkOut).length,
      estimatedPrice: null,
      currency: getDefaultBookingCurrency_(),
      remainingMin: null,
      roomTypeId: roomTypeId,
      roomTypeName: roomTypeName,
      guests: Number(validated.guests || 1),
      message: String(error.message || error),
      missingRateDates: [],
      pricingNotes: [],
      perNight: []
    };
  }
}

function buildAcknowledgementEmailSubject_(requestId) {
  return 'We received your enquiry | Roza\'s Guest House | ' + requestId;
}

function buildAcknowledgementEmailBody_(details) {
  return [
    'Hello,',
    '',
    'We received your enquiry for Roza\'s Guest House.',
    '',
    'Reference number: ' + details.requestId,
    'Check-in: ' + formatDateForGuest_(details.checkIn),
    'Check-out: ' + formatDateForGuest_(details.checkOut),
    'Room type: ' + details.roomTypeName,
    'Preferred bed setup: ' + (details.bedSetup || 'Best available'),
    'Guests: ' + details.guests,
    '',
    'This is an acknowledgement only. Your stay is not confirmed yet.',
    'Roza will review your enquiry and reply as soon as possible.',
    '',
    'Thank you,',
    'Roza\'s Guest House'
  ].join('\n');
}

function getInitialAcknowledgementStatus_(guestEmail) {
  if (!guestEmail) return 'not-requested';
  return isValidEmailAddress_(guestEmail) ? 'pending' : 'invalid-email';
}

function getInitialAcknowledgementError_(guestEmail) {
  if (!guestEmail || isValidEmailAddress_(guestEmail)) return '';
  return 'Guest email format is invalid.';
}

function buildWebsiteBookingConfirmationEmailSubject_(bookingId) {
  return 'Your booking is confirmed | Roza\'s Guest House | ' + bookingId;
}

function escapeHtmlForEmail_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function(character) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    }[character];
  });
}

function buildBookingValueSummary_(originalValue, bookingCurrency, gbpValue) {
  const reportingCurrency = getReportingCurrency_();
  const normalizedBookingCurrency = normalizeCurrencyCode_(bookingCurrency) || reportingCurrency;
  const originalAmount = roundCurrency_(originalValue || 0);
  const reportingAmount = roundCurrency_(gbpValue || 0);

  if (normalizedBookingCurrency !== reportingCurrency && originalAmount) {
    return formatMoneyValue_(originalAmount, normalizedBookingCurrency) + ' / ' + formatMoneyValue_(reportingAmount, reportingCurrency);
  }

  return formatMoneyValue_(reportingAmount || originalAmount, reportingCurrency);
}

function buildWebsiteBookingConfirmationEmailHtml_(details) {
  const propertyName = escapeHtmlForEmail_(details.propertyName || DEFAULT_PROPERTY_NAME);
  const guestName = escapeHtmlForEmail_(details.guestName || 'Guest');
  const bookingId = escapeHtmlForEmail_(details.bookingId || '');
  const roomTypeName = escapeHtmlForEmail_(details.roomTypeName || '');
  const checkIn = escapeHtmlForEmail_(formatDateForGuest_(details.checkIn));
  const checkOut = escapeHtmlForEmail_(formatDateForGuest_(details.checkOut));
  const guests = escapeHtmlForEmail_(details.guests);
  const bedSetup = escapeHtmlForEmail_(details.bedSetup || 'Best available');
  const bookingValue = escapeHtmlForEmail_(buildBookingValueSummary_(details.bookingValueOriginal || details.bookingValue, details.bookingCurrency || details.currency, details.bookingValueGbp || details.bookingValue));
  const balanceDue = escapeHtmlForEmail_(formatMoneyValue_(details.balanceDue, details.reportingCurrency || getReportingCurrency_()));
  const propertyAddress = escapeHtmlForEmail_(details.propertyAddress || DEFAULT_PROPERTY_ADDRESS);
  const propertyWhatsapp = String(details.propertyWhatsapp || details.propertyPhone || DEFAULT_PROPERTY_PHONE).trim();
  const propertyPhone = String(details.propertyPhone || details.propertyWhatsapp || DEFAULT_PROPERTY_PHONE).trim();
  const propertyEmail = String(details.propertyEmail || DEFAULT_PROPERTY_EMAIL).trim();
  const whatsappDigits = propertyWhatsapp.replace(/[^\d]/g, '');
  const phoneHref = propertyPhone.replace(/[^\d+]/g, '');
  const emailHref = escapeHtmlForEmail_(propertyEmail);
  const notes = String(details.notes || '').trim();
  const notesHtml = notes
    ? [
        '<tr>',
        '  <td style="padding: 0 0 18px 0;">',
        '    <div style="padding: 14px 16px; border-radius: 16px; background: #f8f2eb; border: 1px solid #eadfce; color: #4c4338; font-size: 14px; line-height: 1.6;">',
        '      <div style="font-weight: 700; color: #2d3a2f; margin-bottom: 6px;">Your note</div>',
        '      <div>' + escapeHtmlForEmail_(notes).replace(/\r?\n/g, '<br>') + '</div>',
        '    </div>',
        '  </td>',
        '</tr>'
      ].join('')
    : '';

  return [
    '<!DOCTYPE html>',
    '<html>',
    '<body style="margin:0; padding:0; background:#f4ede5; color:#2d3a2f; font-family:Arial, Helvetica, sans-serif;">',
    '  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; background:#f4ede5;">',
    '    <tr>',
    '      <td align="center" style="padding: 28px 14px;">',
    '        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%; max-width:640px; border-collapse:collapse;">',
    '          <tr>',
    '            <td style="padding: 0 0 18px 0; text-align:center; color:#7b6854; font-size:12px; letter-spacing:0.18em; text-transform:uppercase;">Roza\'s Guest House</td>',
    '          </tr>',
    '          <tr>',
    '            <td style="background:#fffaf5; border:1px solid #eadfce; border-radius:28px; box-shadow:0 18px 46px rgba(28,38,31,0.08); overflow:hidden;">',
    '              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse;">',
    '                <tr>',
    '                  <td style="padding: 26px 28px 10px 28px; background:linear-gradient(180deg, #fcf5ed 0%, #fffaf5 100%);">',
    '                    <div style="font-size:12px; letter-spacing:0.14em; text-transform:uppercase; color:#a27d5e; font-weight:700;">Your booking is confirmed</div>',
    '                    <div style="padding-top:8px; font-family:Georgia, Times New Roman, serif; font-size:32px; line-height:1.2; color:#243128; font-weight:700;">We look forward to welcoming you</div>',
    '                    <div style="padding-top:10px; font-size:15px; line-height:1.7; color:#5f6267;">Hello ' + guestName + ', thank you for booking with ' + propertyName + '. Your stay is now confirmed and your confirmation number is below.</div>',
    '                  </td>',
    '                </tr>',
    '                <tr>',
    '                  <td style="padding: 0 28px 22px 28px;">',
    '                    <div style="padding:16px 18px; border-radius:20px; background:#2d3a2f; color:#fdf8f2;">',
    '                      <div style="font-size:12px; letter-spacing:0.12em; text-transform:uppercase; color:#d9c6b2;">Confirmation number</div>',
    '                      <div style="padding-top:6px; font-size:24px; font-weight:700; letter-spacing:0.02em;">' + bookingId + '</div>',
    '                    </div>',
    '                  </td>',
    '                </tr>',
    '                <tr>',
    '                  <td style="padding: 0 28px 22px 28px;">',
    '                    <div style="border:1px solid #eadfce; border-radius:20px; overflow:hidden; background:#ffffff;">',
    '                      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse;">',
    '                        <tr>',
    '                          <td colspan="2" style="padding: 16px 18px; background:#f8f2eb; color:#243128; font-size:15px; font-weight:700;">Stay details</td>',
    '                        </tr>',
    '                        <tr><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#7b6854; font-size:13px; width:38%;">Room</td><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#243128; font-size:14px; font-weight:600;">' + roomTypeName + '</td></tr>',
    '                        <tr><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#7b6854; font-size:13px;">Check-in</td><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#243128; font-size:14px; font-weight:600;">' + checkIn + '</td></tr>',
    '                        <tr><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#7b6854; font-size:13px;">Check-out</td><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#243128; font-size:14px; font-weight:600;">' + checkOut + '</td></tr>',
    '                        <tr><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#7b6854; font-size:13px;">Guests</td><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#243128; font-size:14px; font-weight:600;">' + guests + '</td></tr>',
    '                        <tr><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#7b6854; font-size:13px;">Bed setup</td><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#243128; font-size:14px; font-weight:600;">' + bedSetup + '</td></tr>',
    '                        <tr><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#7b6854; font-size:13px;">Booking value</td><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#243128; font-size:14px; font-weight:600;">' + bookingValue + '</td></tr>',
    '                        <tr><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#7b6854; font-size:13px;">Balance due</td><td style="padding: 13px 18px; border-top:1px solid #f0e6d9; color:#243128; font-size:14px; font-weight:600;">' + balanceDue + '</td></tr>',
    '                      </table>',
    '                    </div>',
    '                  </td>',
    '                </tr>',
    notesHtml,
    '                <tr>',
    '                  <td style="padding: 0 28px 22px 28px;">',
    '                    <div style="padding:16px 18px; border-radius:20px; background:#f7f1e8; border:1px solid #eadfce;">',
    '                      <div style="font-size:15px; font-weight:700; color:#243128; margin-bottom:8px;">Need help before arrival?</div>',
    '                      <div style="font-size:14px; line-height:1.7; color:#5f6267;">If your arrival time changes, or if you need help with transport or directions, please message us on WhatsApp or reply by email. We will be happy to help.</div>',
    '                      <div style="font-size:14px; line-height:1.7; color:#5f6267; margin-top:8px;">No online payment has been taken yet for this booking.</div>',
    '                    </div>',
    '                  </td>',
    '                </tr>',
    '                <tr>',
    '                  <td style="padding: 0 28px 24px 28px;">',
    '                    <div style="border:1px solid #eadfce; border-radius:20px; background:#ffffff; overflow:hidden;">',
    '                      <div style="padding: 16px 18px; background:#f8f2eb; color:#243128; font-size:15px; font-weight:700;">Contact</div>',
    '                      <div style="padding: 16px 18px 18px 18px;">',
    '                        <div style="font-size:14px; line-height:1.7; color:#243128; font-weight:700;">' + propertyName + '</div>',
    '                        <div style="font-size:14px; line-height:1.7; color:#5f6267;">' + propertyAddress + '</div>',
    '                        <div style="font-size:14px; line-height:1.7; color:#5f6267;">WhatsApp / Phone: ' + escapeHtmlForEmail_(propertyWhatsapp) + '</div>',
    '                        <div style="font-size:14px; line-height:1.7; color:#5f6267;">Email: ' + escapeHtmlForEmail_(propertyEmail) + '</div>',
    '                        <div style="padding-top:14px;">',
    '                          <a href="https://wa.me/' + whatsappDigits + '" style="display:inline-block; margin:0 10px 10px 0; padding:12px 18px; border-radius:999px; background:#1faa5a; color:#ffffff; text-decoration:none; font-size:14px; font-weight:700;">WhatsApp us</a>',
    '                          <a href="mailto:' + emailHref + '" style="display:inline-block; margin:0 10px 10px 0; padding:12px 18px; border-radius:999px; background:#243128; color:#ffffff; text-decoration:none; font-size:14px; font-weight:700;">Email us</a>',
    '                          <a href="tel:' + escapeHtmlForEmail_(phoneHref) + '" style="display:inline-block; margin:0 10px 10px 0; padding:12px 18px; border-radius:999px; background:#f4ede5; color:#243128; text-decoration:none; font-size:14px; font-weight:700; border:1px solid #eadfce;">Call us</a>',
    '                        </div>',
    '                      </div>',
    '                    </div>',
    '                  </td>',
    '                </tr>',
    '                <tr>',
    '                  <td style="padding: 0 28px 30px 28px; font-size:14px; line-height:1.7; color:#5f6267;">Warm regards,<br><span style="color:#243128; font-weight:700;">Roza\'s Guest House family</span></td>',
    '                </tr>',
    '              </table>',
    '            </td>',
    '          </tr>',
    '        </table>',
    '      </td>',
    '    </tr>',
    '  </table>',
    '</body>',
    '</html>'
  ].join('');
}

function buildWebsiteBookingConfirmationEmailBody_(details) {
  return [
    'Hello ' + details.guestName + ',',
    '',
    'Your booking is confirmed at ' + details.propertyName + '.',
    '',
    'Confirmation number: ' + details.bookingId,
    '',
    'Stay details',
    'Room type: ' + details.roomTypeName,
    'Check-in: ' + formatDateForGuest_(details.checkIn),
    'Check-out: ' + formatDateForGuest_(details.checkOut),
    'Guests: ' + details.guests,
    'Preferred bed setup: ' + (details.bedSetup || 'Best available'),
    'Booking value: ' + buildBookingValueSummary_(details.bookingValueOriginal || details.bookingValue, details.bookingCurrency || details.currency, details.bookingValueGbp || details.bookingValue),
    'Payment status: ' + details.paymentStatus,
    'Balance due: ' + formatMoneyValue_(details.balanceDue, details.reportingCurrency || getReportingCurrency_()),
    '',
    'Property: ' + details.propertyName,
    'Address: ' + details.propertyAddress,
    'WhatsApp / phone: ' + details.propertyWhatsapp,
    'Email: ' + details.propertyEmail,
    details.notes ? 'Special notes: ' + details.notes : '',
    '',
    'No online payment has been taken for this booking yet.',
    'If your arrival time changes, or if you need help with transport or directions, please contact us on WhatsApp or by email.',
    '',
    'Contact',
    details.propertyName,
    details.propertyAddress,
    'WhatsApp / phone: ' + details.propertyWhatsapp,
    'Email: ' + details.propertyEmail,
    '',
    'Warm regards,',
    'Roza\'s Guest House family'
  ].filter(Boolean).join('\n');
}

function buildWebsiteBookingNotificationSubject_(bookingId) {
  return 'New direct website booking | ' + bookingId;
}

function buildWebsiteBookingNotificationBody_(details) {
  return [
    'A new direct website booking has been created.',
    '',
    'Confirmation number: ' + details.bookingId,
    'Guest: ' + details.guestName,
    'Email: ' + details.guestEmail,
    'Phone: ' + details.guestPhone,
    'Country: ' + details.country,
    'Room type: ' + details.roomTypeName,
    'Check-in: ' + formatDateForGuest_(details.checkIn),
    'Check-out: ' + formatDateForGuest_(details.checkOut),
    'Guests: ' + details.guests,
    'Preferred bed setup: ' + (details.bedSetup || 'Best available'),
    'Booking value: ' + buildBookingValueSummary_(details.bookingValueOriginal || details.bookingValue, details.bookingCurrency || details.currency, details.bookingValueGbp || details.bookingValue),
    'Payment status: ' + details.paymentStatus,
    'Balance due: ' + formatMoneyValue_(details.balanceDue, details.reportingCurrency || getReportingCurrency_()),
    'Source: Direct Website',
    details.notes ? 'Notes: ' + details.notes : ''
  ].filter(Boolean).join('\n');
}

function formatWebsiteEmailFollowUpTimestamp_(value) {
  const date = normalizeDateTimeInput_(value) || new Date();
  return Utilities.formatDate(date, getScriptTimeZone_(), 'yyyy-MM-dd HH:mm');
}

function buildWebsiteEmailFollowUpNote_(status, details) {
  const info = details || {};
  const parts = [
    'Website email follow-up: ' + String(status || 'pending').trim(),
    'at ' + formatWebsiteEmailFollowUpTimestamp_(info.at || info.timestamp || new Date())
  ];
  if (info.guestStatus) {
    parts.push('guest confirmation=' + String(info.guestStatus || '').trim());
  }
  if (info.internalStatus) {
    parts.push('internal notification=' + String(info.internalStatus || '').trim());
  }
  if (info.message) {
    parts.push(String(info.message || '').trim());
  }
  if (info.error) {
    parts.push('error=' + String(info.error || '').trim());
  }
  return parts.filter(Boolean).join('; ');
}

function cleanWebsiteEmailFollowUpNotes_(notes) {
  return String(notes || '').split('|').map(function(part) {
    return String(part || '').trim();
  }).filter(function(part) {
    if (!part) return false;
    if (/^Website email follow-up:/i.test(part)) return false;
    if (/^Guest confirmation email not sent synchronously/i.test(part)) return false;
    if (/^Internal notification email not sent synchronously/i.test(part)) return false;
    return true;
  }).join(' | ');
}

function setWebsiteEmailFollowUpNote_(existingNotes, status, details) {
  return mergeOperationalNotes_(
    cleanWebsiteEmailFollowUpNotes_(existingNotes),
    buildWebsiteEmailFollowUpNote_(status, details)
  );
}

function parseWebsiteEmailFollowUpTimestamp_(notes) {
  const match = String(notes || '').match(/Website email follow-up:[^|]*?\bat\s+(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})/i);
  if (!match) return null;
  const parsed = new Date(
    Number(match[1]),
    Number(match[2]) - 1,
    Number(match[3]),
    Number(match[4]),
    Number(match[5])
  );
  return isNaN(parsed.getTime()) ? null : parsed;
}

function isWebsiteEmailFollowUpSendingStale_(notes) {
  const timestamp = parseWebsiteEmailFollowUpTimestamp_(notes);
  if (!timestamp) return true;
  return new Date().getTime() - timestamp.getTime() > WEBSITE_EMAIL_FOLLOW_UP_SENDING_STALE_MINUTES * 60 * 1000;
}

function getWebsiteEmailFollowUpStateFromNotes_(notes) {
  const text = String(notes || '');
  if (/Website email follow-up:\s*sent/i.test(text)) return 'sent';
  if (/Website email follow-up:\s*skipped/i.test(text)) return 'skipped';
  if (/Website email follow-up:\s*failed/i.test(text)) return 'failed';
  if (/Website email follow-up:\s*sending/i.test(text)) {
    return isWebsiteEmailFollowUpSendingStale_(text) ? 'pending' : 'sending';
  }
  if (/Website email follow-up:\s*(pending|queued)/i.test(text)) return 'pending';
  if (/Guest confirmation email not sent synchronously|Website post-commit follow-up deferred/i.test(text)) return 'pending';
  return '';
}

function queueWebsiteBookingEmailFollowUp_(bookingId, options) {
  const queueStartedAt = Date.now();
  logTiming_('queueWebsiteBookingEmailFollowUp:start');
  const queuedAt = new Date();
  logTiming_('queueWebsiteBookingEmailFollowUp:permanentTrigger', queueStartedAt);
  logTiming_('queueWebsiteBookingEmailFollowUp:end', queueStartedAt);
  return {
    status: 'queued',
    queuedAt: queuedAt,
    error: ''
  };
}

function markWebsiteEmailFollowUpScheduleFailure_(bookingId, message) {
  const targetBookingId = String(bookingId || '').trim();
  if (!targetBookingId) return;
  try {
    const bookingRef = getBookingRefById_(targetBookingId);
    const existingNotes = String(bookingRef.rowObject && bookingRef.rowObject.internal_notes || '').trim();
    updateObjectRowBulk_(bookingRef.sheet, bookingRef.rowNumber, {
      internal_notes: setWebsiteEmailFollowUpNote_(existingNotes, 'pending', {
        guestStatus: 'pending',
        internalStatus: 'pending',
        message: 'Automatic trigger scheduling failed. Use the Mini PMS menu item Send Pending Website Emails.',
        error: message
      })
    });
  } catch (noteError) {
    // Booking is already committed. Do not throw from recovery-note writing.
  }
}

function buildWebsiteBookingEmailDetailsFromBookingRow_(row) {
  const contactDetails = getWebsiteContactDetails_();
  const defaultBookingCurrency = getDefaultBookingCurrency_();
  const reportingCurrency = getReportingCurrency_();
  const bookingCurrency = getStoredBookingCurrency_(row, defaultBookingCurrency);
  const bookingValueOriginal = getStoredBookingOriginalValue_(row);
  const bookingValueGbp = getOperationalBookingValueGbp_(row, {
    defaultBookingCurrency: defaultBookingCurrency
  });
  const amountPaid = roundCurrency_(row.amount_paid || 0);
  const balanceDue = row.balance_due != null && row.balance_due !== ''
    ? roundCurrency_(row.balance_due)
    : calculateBalanceDue_(bookingValueGbp, amountPaid);

  return {
    bookingId: String(row.booking_id || '').trim(),
    guestEmail: String(row.guest_email || '').trim(),
    guestName: String(row.guest_name || '').trim(),
    guestPhone: String(row.guest_phone || '').trim(),
    country: String(row.country || '').trim(),
    roomTypeName: String(row.room_type_name || row.room_type_id || '').trim(),
    checkIn: row.check_in,
    checkOut: row.check_out,
    guests: Math.max(1, Number(row.guests || Number(row.adults || 0) + Number(row.children || 0) || 1)),
    bedSetup: normalizeBedSetup_(row.bed_setup || '') || 'Best available',
    bookingValue: bookingValueOriginal,
    bookingValueOriginal: bookingValueOriginal,
    bookingValueGbp: bookingValueGbp,
    pricingSource: String(row.pricing_source || '').trim() || 'base_rate',
    pricingReferenceId: String(row.pricing_reference_id || '').trim(),
    balanceDue: balanceDue,
    paymentStatus: normalizePaymentStatus_(row.payment_status, bookingValueGbp, amountPaid),
    currency: bookingCurrency,
    bookingCurrency: bookingCurrency,
    reportingCurrency: reportingCurrency,
    fxRateToGbp: row.fx_rate_to_gbp || (bookingCurrency === reportingCurrency ? 1 : ''),
    notes: String(row.notes || '').trim(),
    propertyName: contactDetails.propertyName,
    propertyAddress: contactDetails.propertyAddress,
    propertyPhone: contactDetails.phone,
    propertyWhatsapp: contactDetails.whatsapp,
    propertyEmail: contactDetails.email,
    notificationEmail: contactDetails.notificationEmail
  };
}

function buildWebsiteBookingEmailDetailsFromCommittedBooking_(bookingResult, bookingContext) {
  const result = bookingResult || {};
  const context = bookingContext || {};
  const contactDetails = getWebsiteContactDetails_();
  const reportingCurrency = result.reportingCurrency || getReportingCurrency_();
  const bookingCurrency = result.bookingCurrency || result.currency || DEFAULT_CURRENCY;
  const bookingValueOriginal = result.bookingValueOriginal != null && result.bookingValueOriginal !== ''
    ? result.bookingValueOriginal
    : result.bookingValue;
  const bookingValueGbp = result.bookingValueGbp != null && result.bookingValueGbp !== ''
    ? result.bookingValueGbp
    : result.bookingValue;
  const amountPaid = 0;
  const balanceDue = result.balanceDue != null && result.balanceDue !== ''
    ? roundCurrency_(result.balanceDue)
    : calculateBalanceDue_(bookingValueGbp, amountPaid);

  return {
    bookingId: String(result.bookingId || '').trim(),
    guestEmail: String(context.guestEmail || '').trim(),
    guestName: String(context.guestName || '').trim(),
    guestPhone: String(context.guestPhone || '').trim(),
    country: String(context.country || '').trim(),
    roomTypeName: String(result.roomTypeName || context.roomTypeName || result.roomTypeId || context.roomTypeId || '').trim(),
    checkIn: result.checkIn || context.checkIn,
    checkOut: result.checkOut || context.checkOut,
    guests: Math.max(1, Number(context.guests || 1)),
    bedSetup: normalizeBedSetup_(context.bedSetup || '') || 'Best available',
    bookingValue: bookingValueOriginal,
    bookingValueOriginal: bookingValueOriginal,
    bookingValueGbp: bookingValueGbp,
    pricingSource: String(result.pricingSource || '').trim() || 'base_rate',
    pricingReferenceId: String(result.pricingReferenceId || '').trim(),
    balanceDue: balanceDue,
    paymentStatus: normalizePaymentStatus_(result.paymentStatus || 'Unpaid', bookingValueGbp, amountPaid),
    currency: bookingCurrency,
    bookingCurrency: bookingCurrency,
    reportingCurrency: reportingCurrency,
    fxRateToGbp: result.fxRateToGbp || (bookingCurrency === reportingCurrency ? 1 : ''),
    notes: String(context.notes || '').trim(),
    propertyName: contactDetails.propertyName,
    propertyAddress: contactDetails.propertyAddress,
    propertyPhone: contactDetails.phone,
    propertyWhatsapp: contactDetails.whatsapp,
    propertyEmail: contactDetails.email,
    notificationEmail: contactDetails.notificationEmail
  };
}

function getPendingWebsiteBookingEmailRows_(sheet, options) {
  const config = options || {};
  const retryFailed = toBoolean_(config.retryFailed || config.retry_failed);
  const headerMap = getHeaderMap_(sheet);
  const rows = getSheetRowsForHeaders_(sheet, headerMap, [
    'booking_id',
    'created_at',
    'source',
    'guest_name',
    'guest_phone',
    'guest_email',
    'country',
    'check_in',
    'check_out',
    'room_type_id',
    'room_type_name',
    'bed_setup',
    'adults',
    'children',
    'guests',
    'booking_value',
    'booking_value_original',
    'booking_currency',
    'fx_rate_to_gbp',
    'booking_value_gbp',
    'pricing_source',
    'pricing_reference_id',
    'amount_paid',
    'payment_status',
    'balance_due',
    'status',
    'notes',
    'internal_notes'
  ]);

  return rows.map(function(row, index) {
    row.rowNumber = index + 2;
    return row;
  }).filter(function(row) {
    const bookingId = String(row.booking_id || '').trim();
    if (!bookingId) return false;
    if (normalizeBookingSource_(row.source || '') !== 'Direct Website') return false;
    const bookingStatus = normalizeBookingStatus_(row.status);
    if (bookingStatus === BOOKING_STATUS_CANCELLED || bookingStatus === BOOKING_STATUS_NO_SHOW) return false;
    const followUpState = getWebsiteEmailFollowUpStateFromNotes_(row.internal_notes);
    return followUpState === 'pending' || (retryFailed && followUpState === 'failed');
  });
}

function sendWebsiteBookingEmailFollowUp(input) {
  const payload = input || {};
  const bookingId = String(payload.booking_id || payload.bookingId || '').trim();
  const guestEmail = String(payload.guest_email || payload.guestEmail || '').trim().toLowerCase();
  if (!bookingId) throw new Error('Booking ID is required for email follow-up.');
  if (!guestEmail || !isValidEmailAddress_(guestEmail)) throw new Error('A valid guest email is required for email follow-up.');

  const bookingRef = getBookingRefById_(bookingId);
  return sendWebsiteBookingEmailFollowUpForRow_(bookingRef.sheet, bookingRef.rowNumber, {
    bookingId: bookingId,
    guestEmail: guestEmail,
    retryFailed: true,
    source: 'public-follow-up'
  });
}

function sendWebsiteBookingEmailFollowUpSafely_(bookingResult, guestEmail, bookingContext) {
  const bookingId = String(bookingResult && bookingResult.bookingId || '').trim();
  const rowNumber = Number(bookingResult && bookingResult.rowNumber || 0);
  const expectedGuestEmail = String(guestEmail || '').trim().toLowerCase();
  if (!bookingId || !rowNumber) {
    return {
      ok: false,
      bookingId: bookingId,
      status: 'failed',
      guestStatus: 'failed',
      internalStatus: 'failed',
      error: 'Booking email follow-up could not locate the committed booking row.'
    };
  }

  try {
    if (bookingContext) {
      return deliverCommittedWebsiteBookingEmailFollowUp_(bookingResult, bookingContext);
    }
    const sheet = getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.BOOKINGS);
    return sendWebsiteBookingEmailFollowUpForRow_(sheet, rowNumber, {
      bookingId: bookingId,
      guestEmail: expectedGuestEmail,
      retryFailed: true,
      source: 'immediate-post-commit'
    });
  } catch (error) {
    const message = String(error && error.message || error || 'Booking email follow-up failed.');
    try {
      const bookingRef = getBookingRefById_(bookingId);
      const existingNotes = String(bookingRef.rowObject && bookingRef.rowObject.internal_notes || '').trim();
      updateObjectRowBulk_(bookingRef.sheet, bookingRef.rowNumber, {
        internal_notes: setWebsiteEmailFollowUpNote_(existingNotes, 'failed', {
          guestStatus: 'failed',
          internalStatus: 'failed',
          message: 'Immediate email follow-up failed after booking commit; staff retry is needed.',
          error: message
        })
      });
    } catch (noteError) {
      // Booking is already committed. Do not throw from recovery-note writing.
    }
    return {
      ok: false,
      bookingId: bookingId,
      status: 'failed',
      guestStatus: 'failed',
      internalStatus: 'failed',
      error: message,
      guestError: message,
      internalError: message
    };
  }
}

function updateCellValueByHeader_(sheet, rowNumber, headerName, updater) {
  const headerMap = getHeaderMap_(sheet);
  const columnIndex = headerMap[String(headerName || '').trim()];
  if (columnIndex == null) return false;
  const range = sheet.getRange(rowNumber, columnIndex + 1);
  range.setValue(updater(range.getValue()));
  return true;
}

function deliverCommittedWebsiteBookingEmailFollowUp_(bookingResult, bookingContext) {
  const result = bookingResult || {};
  const bookingId = String(result.bookingId || '').trim();
  const rowNumber = Number(result.rowNumber || 0);
  const details = buildWebsiteBookingEmailDetailsFromCommittedBooking_(result, bookingContext);
  const guestResult = sendWebsiteBookingConfirmationEmail_(details);
  const internalResult = sendWebsiteBookingNotificationEmail_(details);
  const guestFailed = ['failed', 'invalid-email'].indexOf(String(guestResult.status || '')) !== -1;
  const internalFailed = ['failed', 'invalid-email'].indexOf(String(internalResult.status || '')) !== -1;
  const hasFailure = guestFailed || internalFailed;
  const status = hasFailure ? 'failed' : 'sent';
  const message = hasFailure
    ? 'Email follow-up attempted after booking commit; staff review may be needed.'
    : 'Guest confirmation and internal notification follow-up completed after booking commit.';
  const error = [guestResult.error, internalResult.error].filter(Boolean).join(' | ');

  try {
    const sheet = getSheetOrThrow_(getSpreadsheet_(), SHEET_NAMES.BOOKINGS);
    updateCellValueByHeader_(sheet, rowNumber, 'internal_notes', function(latestNotes) {
      return setWebsiteEmailFollowUpNote_(latestNotes || result.internalNotes || '', status, {
        guestStatus: guestResult.status,
        internalStatus: internalResult.status,
        message: message,
        error: error
      });
    });
  } catch (noteError) {
    // Booking is already committed and email was attempted. Do not fail the guest confirmation for note writing.
  }

  return {
    ok: !hasFailure,
    bookingId: bookingId,
    status: status,
    guestStatus: guestResult.status,
    internalStatus: internalResult.status,
    guestSentAt: guestResult.sentAt || '',
    internalSentAt: internalResult.sentAt || '',
    guestError: guestResult.error || '',
    internalError: internalResult.error || '',
    error: error
  };
}

function sendWebsiteBookingEmailFollowUpForRow_(sheet, rowNumber, options) {
  const config = options || {};
  const bookingId = String(config.bookingId || config.booking_id || '').trim();
  const expectedGuestEmail = String(config.guestEmail || config.guest_email || '').trim().toLowerCase();
  const retryFailed = toBoolean_(config.retryFailed || config.retry_failed);
  const lock = LockService.getScriptLock();
  let claimedRow = null;

  lock.waitLock(30000);
  try {
    const latestRow = getRowObjectByNumber_(sheet, rowNumber);
    if (!latestRow || String(latestRow.booking_id || '').trim() !== bookingId) {
      return {
        ok: false,
        status: 'not-found',
        bookingId: bookingId,
        error: 'Booking row could not be found for email follow-up.'
      };
    }
    if (normalizeBookingSource_(latestRow.source || '') !== 'Direct Website') {
      return {
        ok: false,
        status: 'not-website-booking',
        bookingId: bookingId,
        error: 'Email follow-up is only available for direct website bookings.'
      };
    }
    if (expectedGuestEmail && String(latestRow.guest_email || '').trim().toLowerCase() !== expectedGuestEmail) {
      return {
        ok: false,
        status: 'email-mismatch',
        bookingId: bookingId,
        error: 'Guest email does not match the booking.'
      };
    }
    const bookingStatus = normalizeBookingStatus_(latestRow.status);
    if (bookingStatus === BOOKING_STATUS_CANCELLED || bookingStatus === BOOKING_STATUS_NO_SHOW) {
      return {
        ok: false,
        status: 'booking-not-active',
        bookingId: bookingId,
        error: 'Email follow-up was skipped because this booking is not active.'
      };
    }

    const followUpState = getWebsiteEmailFollowUpStateFromNotes_(latestRow.internal_notes);
    if (followUpState === 'sent' || followUpState === 'skipped') {
      return {
        ok: true,
        status: followUpState,
        bookingId: bookingId,
        message: 'Website booking email follow-up is already complete.'
      };
    }
    if (followUpState !== 'pending' && !(retryFailed && followUpState === 'failed')) {
      return {
        ok: true,
        status: followUpState || 'not-pending',
        bookingId: bookingId,
        message: 'Website booking email follow-up is not currently pending.'
      };
    }

    const claimNotes = setWebsiteEmailFollowUpNote_(latestRow.internal_notes, 'sending', {
      guestStatus: 'pending',
      internalStatus: 'pending',
      message: 'Email follow-up claimed for immediate delivery after booking confirmation.'
    });
    updateObjectRowBulk_(sheet, rowNumber, {
      internal_notes: claimNotes
    });
    claimedRow = Object.assign({}, latestRow, {
      rowNumber: rowNumber,
      internal_notes: claimNotes
    });
  } finally {
    lock.releaseLock();
  }

  return deliverClaimedWebsiteBookingEmailFollowUp_(sheet, claimedRow);
}

function deliverClaimedWebsiteBookingEmailFollowUp_(sheet, row) {
  const bookingId = String(row && row.booking_id || '').trim();
  const details = buildWebsiteBookingEmailDetailsFromBookingRow_(row);
  const guestResult = sendWebsiteBookingConfirmationEmail_(details);
  const internalResult = sendWebsiteBookingNotificationEmail_(details);
  const guestFailed = ['failed', 'invalid-email'].indexOf(String(guestResult.status || '')) !== -1;
  const internalFailed = ['failed', 'invalid-email'].indexOf(String(internalResult.status || '')) !== -1;
  const hasFailure = guestFailed || internalFailed;
  const status = hasFailure ? 'failed' : 'sent';
  const message = hasFailure
    ? 'Email follow-up attempted after booking commit; staff review may be needed.'
    : 'Guest confirmation and internal notification follow-up completed after booking commit.';
  const error = [guestResult.error, internalResult.error].filter(Boolean).join(' | ');
  let latestNotes = row.internal_notes;

  try {
    const latestRow = getRowObjectByNumber_(sheet, row.rowNumber);
    if (latestRow && String(latestRow.booking_id || '').trim() === bookingId) {
      latestNotes = latestRow.internal_notes;
    }
  } catch (readError) {
    latestNotes = row.internal_notes;
  }

  updateObjectRowBulk_(sheet, row.rowNumber, {
    internal_notes: setWebsiteEmailFollowUpNote_(latestNotes, status, {
      guestStatus: guestResult.status,
      internalStatus: internalResult.status,
      message: message,
      error: error
    })
  });

  return {
    ok: !hasFailure,
    bookingId: bookingId,
    status: status,
    guestStatus: guestResult.status,
    internalStatus: internalResult.status,
    guestSentAt: guestResult.sentAt || '',
    internalSentAt: internalResult.sentAt || '',
    guestError: guestResult.error || '',
    internalError: internalResult.error || '',
    error: error
  };
}

function processPendingWebsiteBookingEmails(options) {
  try {
    return processPendingWebsiteBookingEmailsCore_(options);
  } catch (error) {
    sendSilentFailureAlert_('Website email pending sender', error);
    throw error;
  }
}

function processPendingWebsiteBookingEmailsCore_(options) {
  const config = options || {};
  const batchLimit = Math.max(1, Math.min(25, Number(config.limit || config.max || WEBSITE_EMAIL_FOLLOW_UP_BATCH_LIMIT)));
  const retryFailed = toBoolean_(config.retryFailed || config.retry_failed);
  const result = {
    ok: true,
    processed: 0,
    sent: 0,
    failed: 0,
    skipped: 0,
    remaining: 0,
    rows: []
  };
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  const lock = LockService.getScriptLock();
  const claimedRows = [];
  const pendingRows = getPendingWebsiteBookingEmailRows_(sheet, config);
  const rowsToClaim = pendingRows.slice(0, batchLimit);

  lock.waitLock(30000);
  try {
    rowsToClaim.forEach(function(row) {
      const latestRow = getRowObjectByNumber_(sheet, row.rowNumber);
      if (!latestRow || String(latestRow.booking_id || '').trim() !== String(row.booking_id || '').trim()) return;
      if (normalizeBookingSource_(latestRow.source || '') !== 'Direct Website') return;
      const bookingStatus = normalizeBookingStatus_(latestRow.status);
      if (bookingStatus === BOOKING_STATUS_CANCELLED || bookingStatus === BOOKING_STATUS_NO_SHOW) return;
      const followUpState = getWebsiteEmailFollowUpStateFromNotes_(latestRow.internal_notes);
      if (followUpState !== 'pending' && !(retryFailed && followUpState === 'failed')) return;

      const claimNotes = setWebsiteEmailFollowUpNote_(latestRow.internal_notes, 'sending', {
        guestStatus: 'pending',
        internalStatus: 'pending',
        message: 'Email follow-up claimed for delivery after booking commit.'
      });
      updateObjectRowBulk_(sheet, row.rowNumber, {
        internal_notes: claimNotes
      });
      claimedRows.push(Object.assign({}, row, latestRow, {
        rowNumber: row.rowNumber,
        internal_notes: claimNotes
      }));
    });

    result.remaining = Math.max(0, pendingRows.length - claimedRows.length);
  } finally {
    lock.releaseLock();
  }

  claimedRows.forEach(function(row) {
    const bookingId = String(row.booking_id || '').trim();
    const details = buildWebsiteBookingEmailDetailsFromBookingRow_(row);
    const guestResult = sendWebsiteBookingConfirmationEmail_(details);
    const internalResult = sendWebsiteBookingNotificationEmail_(details);
    const guestFailed = ['failed', 'invalid-email'].indexOf(String(guestResult.status || '')) !== -1;
    const internalFailed = ['failed', 'invalid-email'].indexOf(String(internalResult.status || '')) !== -1;
    const hasFailure = guestFailed || internalFailed;
    const status = hasFailure ? 'failed' : 'sent';
    const message = hasFailure
      ? 'Email follow-up attempted after booking commit; staff review may be needed.'
      : 'Guest confirmation and internal notification follow-up completed after booking commit.';
    const error = [guestResult.error, internalResult.error].filter(Boolean).join(' | ');
    let latestNotes = row.internal_notes;

    try {
      const latestRow = getRowObjectByNumber_(sheet, row.rowNumber);
      if (latestRow && String(latestRow.booking_id || '').trim() === bookingId) {
        latestNotes = latestRow.internal_notes;
      }
    } catch (readError) {
      latestNotes = row.internal_notes;
    }

    updateObjectRowBulk_(sheet, row.rowNumber, {
      internal_notes: setWebsiteEmailFollowUpNote_(latestNotes, status, {
        guestStatus: guestResult.status,
        internalStatus: internalResult.status,
        message: message,
        error: error
      })
    });

    result.processed += 1;
    if (hasFailure) {
      result.failed += 1;
    } else {
      result.sent += 1;
    }
    result.rows.push({
      bookingId: bookingId,
      status: status,
      guestStatus: guestResult.status,
      internalStatus: internalResult.status,
      error: error
    });
  });

  if (result.remaining > 0) {
    queueWebsiteBookingEmailFollowUp_('', { force: true });
  }

  return result;
}

function adminSendPendingWebsiteBookingEmails(input) {
  return processPendingWebsiteBookingEmails(input || {});
}

function retryFailedWebsiteBookingEmails() {
  return processPendingWebsiteBookingEmails({ retryFailed: true });
}

function adminRetryFailedWebsiteBookingEmails(input) {
  return processPendingWebsiteBookingEmails(Object.assign({ retryFailed: true }, input || {}));
}

function sendWebsiteBookingConfirmationEmail_(details) {
  if (!details.guestEmail) {
    return {
      status: 'not-requested',
      sentAt: '',
      error: ''
    };
  }

  if (!isValidEmailAddress_(details.guestEmail)) {
    return {
      status: 'invalid-email',
      sentAt: '',
      error: 'Guest email format is invalid.'
    };
  }

  const senderEmail = DEFAULT_PROPERTY_EMAIL;
  const senderName = DEFAULT_PROPERTY_NAME;
  const subject = buildWebsiteBookingConfirmationEmailSubject_(details.bookingId);
  const body = buildWebsiteBookingConfirmationEmailBody_(details);
  const htmlBody = buildWebsiteBookingConfirmationEmailHtml_(details);
  let aliasFound = false;

  try {
    try {
      aliasFound = GmailApp.getAliases().map(function(alias) {
        return String(alias || '').trim().toLowerCase();
      }).indexOf(String(senderEmail || '').trim().toLowerCase()) !== -1;
    } catch (aliasError) {
      aliasFound = false;
    }

    console.log('Guest confirmation sender alias found: ' + (aliasFound ? 'yes' : 'no'));
    console.log('Guest confirmation GmailApp alias send attempted: yes');

    GmailApp.sendEmail(details.guestEmail, subject, body, {
      from: senderEmail,
      name: senderName,
      replyTo: senderEmail,
      htmlBody: htmlBody
    });

    console.log('Guest confirmation GmailApp alias send succeeded: yes');
    console.log('Guest confirmation MailApp fallback used: no');

    return {
      status: 'sent',
      sentAt: new Date(),
      error: ''
    };
  } catch (gmailError) {
    console.log('Guest confirmation GmailApp alias send succeeded: no');
    console.log('Guest confirmation MailApp fallback used: yes');

    try {
      MailApp.sendEmail({
        to: details.guestEmail,
        subject: subject,
        body: body,
        htmlBody: htmlBody,
        replyTo: senderEmail,
        name: senderName
      });

      return {
        status: 'sent',
        sentAt: new Date(),
        error: ''
      };
    } catch (fallbackError) {
      return {
        status: 'failed',
        sentAt: '',
        error: 'GmailApp alias send failed: ' + String(gmailError.message || gmailError) + ' | MailApp fallback failed: ' + String(fallbackError.message || fallbackError)
      };
    }
  }
}

function sendWebsiteBookingNotificationEmail_(details) {
  const notificationEmail = String(details.notificationEmail || '').trim();
  if (!notificationEmail) {
    return {
      status: 'not-configured',
      sentAt: '',
      error: ''
    };
  }

  if (!isValidEmailAddress_(notificationEmail)) {
    return {
      status: 'invalid-email',
      sentAt: '',
      error: 'Notification email format is invalid.'
    };
  }

  try {
    MailApp.sendEmail({
      to: notificationEmail,
      replyTo: details.guestEmail || '',
      subject: buildWebsiteBookingNotificationSubject_(details.bookingId),
      body: buildWebsiteBookingNotificationBody_(details),
      name: details.propertyName || DEFAULT_PROPERTY_NAME
    });

    return {
      status: 'sent',
      sentAt: new Date(),
      error: ''
    };
  } catch (error) {
    return {
      status: 'failed',
      sentAt: '',
      error: String(error.message || error)
    };
  }
}

function sendAcknowledgementEmail_(details) {
  if (!details.guestEmail) {
    return {
      status: 'not-requested',
      sentAt: '',
      error: ''
    };
  }

  if (!isValidEmailAddress_(details.guestEmail)) {
    return {
      status: 'invalid-email',
      sentAt: '',
      error: 'Guest email format is invalid.'
    };
  }

  try {
    MailApp.sendEmail({
      to: details.guestEmail,
      subject: buildAcknowledgementEmailSubject_(details.requestId),
      body: buildAcknowledgementEmailBody_(details),
      name: 'Roza\'s Guest House'
    });

    return {
      status: 'sent',
      sentAt: new Date(),
      error: ''
    };
  } catch (error) {
    return {
      status: 'failed',
      sentAt: '',
      error: String(error.message || error)
    };
  }
}


function codexSendRozaAliasDiagnosticEmail() {
  const senderEmail = DEFAULT_PROPERTY_EMAIL;
  const targetEmail = String(Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '').trim();
  if (!targetEmail) {
    throw new Error('Could not determine owner email for alias diagnostic.');
  }
  const aliasFound = GmailApp.getAliases().map(function(alias) {
    return String(alias || '').trim().toLowerCase();
  }).indexOf(String(senderEmail || '').trim().toLowerCase()) !== -1;
  console.log('Roza direct alias diagnostic alias found: ' + (aliasFound ? 'yes' : 'no'));
  GmailApp.sendEmail(targetEmail, 'Roza alias sender diagnostic', 'This is a direct Apps Script Gmail alias sender diagnostic sent to the owner only.', {
    from: senderEmail,
    name: DEFAULT_PROPERTY_NAME,
    replyTo: senderEmail
  });
  console.log('Roza direct alias diagnostic GmailApp send succeeded: yes');
  return {
    aliasFound: aliasFound,
    sent: true,
    sentAt: new Date()
  };
}
function runEmailDiagnostics() {
  const spreadsheet = getSpreadsheet_();
  const diagnostics = {
    timestamp: new Date(),
    effectiveEmail: String(Session.getEffectiveUser().getEmail() || '').trim(),
    activeEmail: String(Session.getActiveUser().getEmail() || '').trim(),
    remainingDailyQuota: MailApp.getRemainingDailyQuota(),
    spreadsheetId: spreadsheet.getId(),
    spreadsheetName: spreadsheet.getName()
  };

  Logger.log(JSON.stringify(diagnostics));
  return diagnostics;
}

function sendEmailAuthorisationTest() {
  const diagnostics = runEmailDiagnostics();
  const targetEmail = diagnostics.effectiveEmail || diagnostics.activeEmail;

  if (!targetEmail) {
    throw new Error('Could not determine the Apps Script account email for the mail test.');
  }

  MailApp.sendEmail({
    to: targetEmail,
    subject: 'Roza Booking Engine email test',
    body: [
      'This is a manual test email from the Roza Booking Engine Apps Script.',
      '',
      'If you received this, MailApp is authorised and working for this deployment owner.',
      '',
      'Timestamp: ' + diagnostics.timestamp
    ].join('\n'),
    name: 'Roza\'s Guest House'
  });

  const result = {
    ok: true,
    sentTo: targetEmail,
    sentAt: new Date(),
    remainingDailyQuota: MailApp.getRemainingDailyQuota()
  };

  Logger.log(JSON.stringify(result));
  return result;
}

// -----------------------------------------------------------------------------
// Spreadsheet helpers
// -----------------------------------------------------------------------------

var REQUEST_READ_CACHE_ = null;

function clonePlainObject_(value) {
  return Object.assign({}, value || {});
}

function cloneSheetObjectRows_(rows) {
  return (rows || []).map(function(row) {
    return clonePlainObject_(row);
  });
}

function withRequestReadCache_(callback) {
  const previousCache = REQUEST_READ_CACHE_;
  REQUEST_READ_CACHE_ = {
    sheets: {},
    settingsLoaded: false,
    settings: {}
  };
  try {
    return callback();
  } finally {
    REQUEST_READ_CACHE_ = previousCache;
  }
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const configuredId = String(props.getProperty('SPREADSHEET_ID') || '').trim();

  if (configuredId) {
    return SpreadsheetApp.openById(configuredId);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  throw new Error('No spreadsheet found. Either bind this script to the sheet or set Script Property SPREADSHEET_ID.');
}

function getSheetOrThrow_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Missing sheet: ' + sheetName);
  }
  return sheet;
}

function getSheetHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    throw new Error('Sheet has no headers: ' + sheet.getName());
  }
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
}

function buildHeaderMapFromHeaders_(headers) {
  const map = {};
  (headers || []).forEach(function(header, index) {
    const key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function getHeaderMap_(sheet) {
  return buildHeaderMapFromHeaders_(getSheetHeaders_(sheet));
}

function ensureSheetHeaders_(sheet, requiredHeaders) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    throw new Error('Sheet has no headers: ' + sheet.getName());
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });

  const missingHeaders = requiredHeaders.filter(function(header) {
    return headers.indexOf(header) === -1;
  });

  if (!missingHeaders.length) {
    return;
  }

  sheet.getRange(1, lastColumn + 1, 1, missingHeaders.length).setValues([missingHeaders]);
}

function getSheetObjects_(sheetName) {
  if (REQUEST_READ_CACHE_ && Object.prototype.hasOwnProperty.call(REQUEST_READ_CACHE_.sheets, sheetName)) {
    return cloneSheetObjectRows_(REQUEST_READ_CACHE_.sheets[sheetName]);
  }

  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    if (REQUEST_READ_CACHE_) {
      REQUEST_READ_CACHE_.sheets[sheetName] = [];
    }
    return [];
  }

  const headers = values[0].map(function(header) {
    return String(header || '').trim();
  });

  const rows = values.slice(1).filter(function(row) {
    return row.some(function(cell) { return cell !== '' && cell !== null; });
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });
    return obj;
  });

  if (REQUEST_READ_CACHE_) {
    REQUEST_READ_CACHE_.sheets[sheetName] = cloneSheetObjectRows_(rows);
  }

  return rows;
}

function getSheetRowsForHeaders_(sheet, headerMap, headerKeys) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const keys = [];
  (headerKeys || []).forEach(function(headerKey) {
    const key = String(headerKey || '').trim();
    if (key && Object.prototype.hasOwnProperty.call(headerMap, key) && keys.indexOf(key) === -1) {
      keys.push(key);
    }
  });
  if (!keys.length) return [];

  const rowCount = lastRow - 1;
  const columns = keys.map(function(key) {
    return {
      key: key,
      index: headerMap[key]
    };
  }).sort(function(a, b) {
    return a.index - b.index;
  });
  const groups = [];
  columns.forEach(function(column) {
    const lastGroup = groups.length ? groups[groups.length - 1] : null;
    if (lastGroup && column.index === lastGroup.endIndex + 1) {
      lastGroup.endIndex = column.index;
      lastGroup.columns.push(column);
    } else {
      groups.push({
        startIndex: column.index,
        endIndex: column.index,
        columns: [column]
      });
    }
  });

  const rows = Array.from({ length: rowCount }, function() { return {}; });
  groups.forEach(function(group) {
    const values = sheet.getRange(2, group.startIndex + 1, rowCount, group.endIndex - group.startIndex + 1).getValues();
    group.columns.forEach(function(column) {
      const offset = column.index - group.startIndex;
      for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
        rows[rowIndex][column.key] = values[rowIndex][offset];
      }
    });
  });
  return rows;
}

function getRowObjectsByRowNumbers_(sheet, rowNumbers, headers) {
  const lastRow = sheet.getLastRow();
  const lastColumn = (headers && headers.length) ? headers.length : sheet.getLastColumn();
  const safeHeaders = headers && headers.length ? headers : getSheetHeaders_(sheet);
  const sorted = [];
  const seen = {};

  (rowNumbers || []).forEach(function(rowNumber) {
    const number = Number(rowNumber || 0);
    if (number < 2 || number > lastRow || seen[number]) return;
    seen[number] = true;
    sorted.push(number);
  });

  sorted.sort(function(a, b) { return a - b; });
  if (!sorted.length || lastColumn < 1) return [];

  const groups = [];
  sorted.forEach(function(rowNumber) {
    const lastGroup = groups.length ? groups[groups.length - 1] : null;
    if (lastGroup && rowNumber === lastGroup.end + 1) {
      lastGroup.end = rowNumber;
    } else {
      groups.push({
        start: rowNumber,
        end: rowNumber
      });
    }
  });

  const rows = [];
  groups.forEach(function(group) {
    const values = sheet.getRange(group.start, 1, group.end - group.start + 1, lastColumn).getValues();
    values.forEach(function(valueRow) {
      const row = {};
      safeHeaders.forEach(function(header, index) {
        row[String(header || '').trim()] = valueRow[index];
      });
      rows.push(row);
    });
  });

  return rows;
}

function getFrontDeskBookingRowsForDate_(selectedDate) {
  const date = normalizeFrontDeskDate_(selectedDate);
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  if (sheet.getLastRow() < 2) return [];

  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const requiredHeaders = ['status', 'check_in', 'check_out'];
  const canUseNarrowScan = requiredHeaders.every(function(header) {
    return Object.prototype.hasOwnProperty.call(headerMap, header);
  });
  if (!canUseNarrowScan) {
    return getSheetObjects_(SHEET_NAMES.BOOKINGS).filter(function(row) {
      return matchesFrontDeskBookingMode_(row, date, 'arrivals') ||
        matchesFrontDeskBookingMode_(row, date, 'departures') ||
        matchesFrontDeskBookingMode_(row, date, 'inHouse');
    });
  }

  const scanRows = getSheetRowsForHeaders_(sheet, headerMap, ['status', 'check_in', 'check_out']);
  const rowNumbers = [];
  scanRows.forEach(function(row, index) {
    if (
      matchesFrontDeskBookingMode_(row, date, 'arrivals') ||
      matchesFrontDeskBookingMode_(row, date, 'departures') ||
      matchesFrontDeskBookingMode_(row, date, 'inHouse')
    ) {
      rowNumbers.push(index + 2);
    }
  });

  return getRowObjectsByRowNumbers_(sheet, rowNumbers, headers);
}

function getBookingRowsOverlappingStay_(checkIn, checkOut) {
  const startDate = normalizeDateInput_(checkIn);
  const endDate = normalizeDateInput_(checkOut);
  if (!startDate || !endDate || endDate.getTime() <= startDate.getTime()) return [];

  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  if (sheet.getLastRow() < 2) return [];

  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const rows = getSheetRowsForHeaders_(sheet, headerMap, [
    'booking_id',
    'status',
    'check_in',
    'check_out',
    'room_type_id',
    'room_identifier',
    'guest_name'
  ]);

  return rows.filter(function(row) {
    const status = normalizeBookingStatus_(row.status);
    if (status !== BOOKING_STATUS_CONFIRMED && status !== BOOKING_STATUS_IN_HOUSE) return false;
    const rowCheckIn = normalizeDateInput_(row.check_in);
    const rowCheckOut = normalizeDateInput_(row.check_out);
    return staysOverlap_(startDate, endDate, rowCheckIn, rowCheckOut);
  });
}

function getGuestHistoryRowsForTargetRows_(targetRows) {
  const targets = targetRows || [];
  const targetKeys = {};
  const targetIds = {};
  targets.forEach(function(row) {
    const key = getGuestHistoryKey_(row);
    const bookingId = String(row && row.booking_id || '').trim();
    if (key) targetKeys[key] = true;
    if (bookingId) targetIds[bookingId] = true;
  });

  if (!Object.keys(targetKeys).length) {
    return cloneSheetObjectRows_(targets);
  }

  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  if (sheet.getLastRow() < 2) return cloneSheetObjectRows_(targets);

  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  const historyHeaders = [
    'booking_id',
    'guest_email',
    'guest_phone',
    'guest_name',
    'check_in',
    'check_out',
    'status',
    'guest_preferences',
    'room_type_id',
    'room_type_name',
    'room_identifier',
    'source',
    'booking_value',
    'booking_value_original',
    'booking_currency',
    'fx_rate_to_gbp',
    'booking_value_gbp',
    'notes'
  ];
  const rows = getSheetRowsForHeaders_(sheet, headerMap, historyHeaders);
  const result = [];
  const seenBookingIds = {};

  rows.forEach(function(row) {
    const key = getGuestHistoryKey_(row);
    if (!key || !targetKeys[key]) return;
    const bookingId = String(row.booking_id || '').trim();
    if (bookingId) seenBookingIds[bookingId] = true;
    result.push(row);
  });

  targets.forEach(function(row) {
    const bookingId = String(row && row.booking_id || '').trim();
    if (bookingId && seenBookingIds[bookingId]) return;
    result.push(clonePlainObject_(row));
  });

  return result;
}

function getBookingRowsForAlertCursor_() {
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  if (sheet.getLastRow() < 2) return [];
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  return getSheetRowsForHeaders_(sheet, headerMap, ['created_at']);
}

function getRecentBookingAlertSourceRows_() {
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  if (sheet.getLastRow() < 2) return [];
  const headers = getSheetHeaders_(sheet);
  const headerMap = buildHeaderMapFromHeaders_(headers);
  return getSheetRowsForHeaders_(sheet, headerMap, [
    'booking_id',
    'created_at',
    'source',
    'guest_name',
    'room_type_name',
    'check_in',
    'check_out',
    'balance_due',
    'status',
    'internal_notes'
  ]);
}

function appendObjectRow_(sheet, headerMap, rowObject) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(function(header) {
    const key = String(header || '').trim();
    return Object.prototype.hasOwnProperty.call(rowObject, key) ? rowObject[key] : '';
  });
  const rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  return rowNumber;
}

function updateObjectRow_(sheet, rowNumber, rowObject) {
  const headerMap = getHeaderMap_(sheet);
  Object.keys(rowObject).forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(headerMap, key)) return;
    sheet.getRange(rowNumber, headerMap[key] + 1).setValue(rowObject[key]);
  });
}

function updateObjectRowBulk_(sheet, rowNumber, rowObject, optionalHeaders) {
  const updateStartedAt = Date.now();
  logTiming_('updateObjectRowBulk_:start');
  const lastColumn = sheet.getLastColumn();
  const headers = optionalHeaders && optionalHeaders.length
    ? optionalHeaders.map(function(header) {
        return String(header || '').trim();
      })
    : sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return String(header || '').trim();
  });
  const candidates = [];
  const updatedFields = [];

  headers.forEach(function(header, index) {
    if (index >= lastColumn) return;
    if (!header || !Object.prototype.hasOwnProperty.call(rowObject, header)) return;
    candidates.push({
      header: header,
      index: index,
      value: rowObject[header]
    });
  });

  const readStartedAt = Date.now();
  const readGroups = [];
  candidates.forEach(function(candidate) {
    const lastGroup = readGroups.length ? readGroups[readGroups.length - 1] : null;
    if (lastGroup && candidate.index === lastGroup.endIndex + 1) {
      lastGroup.endIndex = candidate.index;
      lastGroup.candidates.push(candidate);
    } else {
      readGroups.push({
        startIndex: candidate.index,
        endIndex: candidate.index,
        candidates: [candidate]
      });
    }
  });
  readGroups.forEach(function(group) {
    const values = sheet.getRange(rowNumber, group.startIndex + 1, 1, group.endIndex - group.startIndex + 1).getValues()[0];
    group.candidates.forEach(function(candidate) {
      candidate.currentValue = values[candidate.index - group.startIndex];
    });
  });
  logTiming_('updateObjectRowBulk_:readRow', readStartedAt);

  const computeStartedAt = Date.now();
  const changedColumns = [];
  candidates.forEach(function(candidate) {
    if (candidate.currentValue === candidate.value) return;
    changedColumns.push(candidate);
    updatedFields.push(candidate.header);
  });
  logTiming_('updateObjectRowBulk_:computeChanges', computeStartedAt);

  const writeStartedAt = Date.now();
  if (changedColumns.length) {
    const writeGroups = [];
    changedColumns.forEach(function(candidate) {
      const lastGroup = writeGroups.length ? writeGroups[writeGroups.length - 1] : null;
      if (lastGroup && candidate.index === lastGroup.endIndex + 1) {
        lastGroup.endIndex = candidate.index;
        lastGroup.values.push(candidate.value);
      } else {
        writeGroups.push({
          startIndex: candidate.index,
          endIndex: candidate.index,
          values: [candidate.value]
        });
      }
    });
    writeGroups.forEach(function(group) {
      sheet.getRange(rowNumber, group.startIndex + 1, 1, group.values.length).setValues([group.values]);
    });
  }
  logTiming_('updateObjectRowBulk_:writeGroups', writeStartedAt);

  const result = {
    rowNumber: rowNumber,
    changed: changedColumns.length > 0,
    updatedFields: updatedFields
  };
  logTiming_('updateObjectRowBulk_:return', updateStartedAt);
  return result;
}

function ensureSheetWithHeaders_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    if (headers && headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  } else if (headers && headers.length) {
    if (sheet.getLastColumn() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
      ensureSheetHeaders_(sheet, headers);
    }
  }

  if (sheet.getLastColumn() > 0) {
    styleHeaderRow_(sheet, 1, sheet.getLastColumn());
  }

  return sheet;
}

function logTiming_(label, startedAt) {
  try {
    Logger.log(label + (startedAt ? ' ' + (Date.now() - startedAt) + 'ms' : ''));
  } catch (error) {
    // Timing logs must never affect operational paths.
  }
}

function perfStart_(label) {
  const startedAt = Date.now();
  logTiming_('[PERF] ' + label + ':start');
  return startedAt;
}

function perfStep_(label, startedAt) {
  logTiming_('[PERF] ' + label, startedAt);
}

function getMiniPmsRequiredHeadersForSheet_(sheetName) {
  const availabilityCacheHeaders = ['date', 'room_type_id', 'inventory_total', 'booked_confirmed', 'blocked', 'remaining', 'status'];
  const settingsHeaders = ['key', 'value'];
  const roomTypeHeaders = ['room_type_id', 'room_type_name', 'inventory_total', 'max_guests', 'active'];
  const baseRateRequiredHeaders = ['room_type_id', 'room_type_name', 'base_rate', 'extra_guest_fee', 'active', 'updated_at'];
  const commercialControlRequiredHeaders = ['control_id', 'room_type_id', 'room_type_name', 'rule_type', 'start_date', 'end_date', 'override_price', 'overbooking_allowance', 'active', 'note', 'created_at', 'updated_at'];
  const headerMap = {};
  headerMap[SHEET_NAMES.ROOM_TYPES] = roomTypeHeaders;
  headerMap[SHEET_NAMES.ROOMS] = ROOMS_HEADERS;
  headerMap[SHEET_NAMES.BASE_RATES] = baseRateRequiredHeaders;
  headerMap[SHEET_NAMES.RATES] = RATES_HEADERS;
  headerMap[SHEET_NAMES.COMMERCIAL_CONTROLS] = commercialControlRequiredHeaders;
  headerMap[SHEET_NAMES.EVENT_FLAGS] = EVENT_FLAGS_HEADERS;
  headerMap[SHEET_NAMES.COMPETITOR_TRACKER] = COMPETITOR_TRACKER_HEADERS;
  headerMap[SHEET_NAMES.RECOMMENDATION_ACTION_LOG] = RECOMMENDATION_ACTION_LOG_HEADERS;
  headerMap[SHEET_NAMES.OTA_UPDATE_WORKFLOW] = OTA_UPDATE_WORKFLOW_HEADERS;
  headerMap[SHEET_NAMES.BOOKINGS] = BOOKINGS_HEADERS;
  headerMap[SHEET_NAMES.BLOCKED_DATES] = BLOCKED_DATES_HEADERS;
  headerMap[SHEET_NAMES.AVAILABILITY_CACHE] = availabilityCacheHeaders;
  headerMap[SHEET_NAMES.REQUESTS] = REQUESTS_HEADERS;
  headerMap[SHEET_NAMES.SETTINGS] = settingsHeaders;
  headerMap[SHEET_NAMES.BOOKING_NIGHTS] = BOOKING_NIGHTS_HEADERS;
  headerMap[SHEET_NAMES.DAILY_STATS] = DAILY_STATS_HEADERS;
  headerMap[SHEET_NAMES.OTB_SNAPSHOTS] = OTB_SNAPSHOTS_HEADERS;
  headerMap[SHEET_NAMES.WEBSITE_FEEDBACK] = WEBSITE_FEEDBACK_HEADERS;
  return headerMap[sheetName] || [];
}

function assertSheetHeadersReady_(sheet, requiredHeaders) {
  const headers = getSheetHeaders_(sheet);
  const missingHeaders = (requiredHeaders || []).filter(function(header) {
    return headers.indexOf(header) === -1;
  });
  if (missingHeaders.length) {
    throw new Error('Mini PMS setup is incomplete for sheet ' + sheet.getName() + '. Missing headers: ' + missingHeaders.join(', ') + '. Run setupMiniPmsV1 before using this action.');
  }
}

function assertMiniPmsReady_(spreadsheet, requiredSheets) {
  const sheetNames = requiredSheets && requiredSheets.length ? requiredSheets : [
    SHEET_NAMES.ROOM_TYPES,
    SHEET_NAMES.ROOMS,
    SHEET_NAMES.BASE_RATES,
    SHEET_NAMES.RATES,
    SHEET_NAMES.COMMERCIAL_CONTROLS,
    SHEET_NAMES.BOOKINGS,
    SHEET_NAMES.BLOCKED_DATES,
    SHEET_NAMES.AVAILABILITY_CACHE
  ];
  sheetNames.forEach(function(sheetName) {
    const sheet = getSheetOrThrow_(spreadsheet, sheetName);
    assertSheetHeadersReady_(sheet, getMiniPmsRequiredHeadersForSheet_(sheetName));
  });
  return true;
}

function getAvailabilityOperationSheetNames_() {
  return [
    SHEET_NAMES.ROOM_TYPES,
    SHEET_NAMES.ROOMS,
    SHEET_NAMES.BASE_RATES,
    SHEET_NAMES.RATES,
    SHEET_NAMES.COMMERCIAL_CONTROLS,
    SHEET_NAMES.BOOKINGS,
    SHEET_NAMES.BLOCKED_DATES
  ];
}

function getAvailabilityCacheOperationSheetNames_() {
  return [
    SHEET_NAMES.ROOM_TYPES,
    SHEET_NAMES.ROOMS,
    SHEET_NAMES.COMMERCIAL_CONTROLS,
    SHEET_NAMES.BOOKINGS,
    SHEET_NAMES.BLOCKED_DATES,
    SHEET_NAMES.AVAILABILITY_CACHE
  ];
}

function writeRowUpdatesBatched_(sheet, rowUpdates, columnCount) {
  const updates = (rowUpdates || []).filter(function(update) {
    return update && Number(update.rowNumber || 0) >= 2 && update.values && update.values.length;
  }).sort(function(a, b) {
    return Number(a.rowNumber || 0) - Number(b.rowNumber || 0);
  });

  let groupsWritten = 0;
  let index = 0;
  while (index < updates.length) {
    const first = updates[index];
    const groupRows = [first.values];
    let lastRowNumber = Number(first.rowNumber || 0);
    index += 1;
    while (index < updates.length && Number(updates[index].rowNumber || 0) === lastRowNumber + 1) {
      groupRows.push(updates[index].values);
      lastRowNumber = Number(updates[index].rowNumber || 0);
      index += 1;
    }
    sheet.getRange(Number(first.rowNumber || 0), 1, groupRows.length, columnCount).setValues(groupRows);
    groupsWritten += 1;
  }
  return groupsWritten;
}

function normalizeCellForWriteComparison_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDateKey_(value);
  }
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function rowValuesMatchForWrite_(existingValues, nextValues) {
  if (!existingValues || !nextValues || existingValues.length < nextValues.length) return false;
  for (let index = 0; index < nextValues.length; index++) {
    if (normalizeCellForWriteComparison_(existingValues[index]) !== normalizeCellForWriteComparison_(nextValues[index])) {
      return false;
    }
  }
  return true;
}

function findRowNumberByHeaderValue_(sheet, headerKey, lookupValue) {
  const value = String(lookupValue || '').trim();
  if (!value || sheet.getLastRow() < 2) return 0;

  const headerMap = getHeaderMap_(sheet);
  if (!Object.prototype.hasOwnProperty.call(headerMap, headerKey)) {
    return 0;
  }

  const column = headerMap[headerKey] + 1;
  const lookupRange = sheet.getRange(2, column, sheet.getLastRow() - 1, 1);

  try {
    const match = lookupRange
      .createTextFinder(value)
      .matchEntireCell(true)
      .findNext();
    if (match) {
      return match.getRow();
    }
  } catch (error) {
    // Fall back to an in-memory scan if TextFinder is unavailable in this runtime.
  }

  const values = lookupRange.getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === value) {
      return i + 2;
    }
  }

  return 0;
}

function getBookingRefById_(bookingId) {
  const spreadsheet = getSpreadsheet_();
  const sheet = getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS);
  assertSheetHeadersReady_(sheet, BOOKINGS_HEADERS);
  const rowNumber = findRowNumberByHeaderValue_(sheet, 'booking_id', bookingId);

  if (!rowNumber) {
    throw new Error('Booking not found: ' + bookingId);
  }

  return {
    sheet: sheet,
    rowNumber: rowNumber,
    rowObject: getRowObjectByNumber_(sheet, rowNumber)
  };
}

function getBookingDetailById_(bookingId, options) {
  const config = options || {};
  const bookingRef = getBookingRefById_(bookingId);
  const currentRow = bookingRef.rowObject;
  const defaultBookingCurrency = getDefaultBookingCurrency_();
  let historyIndex = null;
  if (!toBoolean_(config.skipGuestHistory || config.fast || config.minimal)) {
    const bookings = getGuestHistoryRowsForTargetRows_([currentRow]);
    historyIndex = buildGuestHistoryIndex_(bookings, {
      defaultBookingCurrency: defaultBookingCurrency,
      targetBookingIds: [bookingId]
    });
  }
  const roomIndex = buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const detail = mapBookingRowToAdminDisplay_(currentRow, {
    historyIndex: historyIndex,
    roomIndex: roomIndex,
    reportingCurrency: getReportingCurrency_(),
    defaultBookingCurrency: defaultBookingCurrency
  });
  if (toBoolean_(config.skipGuestHistory || config.fast || config.minimal)) {
    delete detail.isRepeatGuest;
    delete detail.repeatStayCount;
    delete detail.lastStayCheckOut;
    delete detail.guestHistory;
    delete detail.lastPreference;
  }
  return detail;
}

function getRowObjectByNumber_(sheet, rowNumber) {
  if (!rowNumber || rowNumber < 2 || rowNumber > sheet.getLastRow()) {
    return null;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getValues()[0];
  const result = {};

  headers.forEach(function(header, index) {
    result[String(header || '').trim()] = values[index];
  });

  return result;
}

// -----------------------------------------------------------------------------
// Utility helpers
// -----------------------------------------------------------------------------

function generateRequestId_(sheet, headerMap, now) {
  if (!Object.prototype.hasOwnProperty.call(headerMap, 'request_id')) {
    throw new Error('Requests sheet is missing required header: request_id');
  }

  const prefix = 'REQ-' + Utilities.formatDate(now, getScriptTimeZone_(), 'yyyyMMdd') + '-';
  const requestIdColumn = headerMap.request_id + 1;
  const lastRow = sheet.getLastRow();
  let nextSequence = 1;

  if (lastRow > 1) {
    const values = sheet.getRange(2, requestIdColumn, lastRow - 1, 1).getValues();
    values.forEach(function(row) {
      const currentValue = String(row[0] || '').trim();
      if (currentValue.indexOf(prefix) !== 0) return;

      const match = currentValue.match(/-(\d+)$/);
      if (!match) return;

      const currentSequence = Number(match[1]);
      if (currentSequence >= nextSequence) {
        nextSequence = currentSequence + 1;
      }
    });
  }

  return prefix + padNumber_(nextSequence, 3);
}

function padNumber_(value, width) {
  const text = String(Math.max(0, Number(value) || 0));
  return text.length >= width ? text : new Array(width - text.length + 1).join('0') + text;
}

function isValidEmailAddress_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function resolveRoomTypeId_(roomType) {
  const raw = String(roomType || '').trim();
  if (!raw) throw new Error('Room type is required.');

  if (ROOM_TYPE_NAME_TO_ID[raw]) {
    return ROOM_TYPE_NAME_TO_ID[raw];
  }

  const knownRoomTypeIds = Object.keys(ROOM_TYPE_NAME_TO_ID).map(function(name) {
    return ROOM_TYPE_NAME_TO_ID[name];
  });
  if (knownRoomTypeIds.indexOf(raw) !== -1) {
    return raw;
  }

  const roomTypeRows = getSheetObjects_(SHEET_NAMES.ROOM_TYPES);
  const directId = roomTypeRows.find(function(row) {
    return String(row.room_type_id || '').trim() === raw;
  });
  if (directId) return raw;

  const byName = roomTypeRows.find(function(row) {
    return String(row.room_type_name || '').trim() === raw;
  });
  if (byName) return String(byName.room_type_id || '').trim();

  throw new Error('Unknown room type: ' + raw);
}

function normalizeDateInput_(value) {
  if (!value && value !== 0) return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return stripTime_(value);
  }

  const text = String(value).trim();
  if (!text) return null;

  // Preferred sheet/site format: YYYY-MM-DD
  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
  }

  const parsed = new Date(text);
  if (isNaN(parsed)) {
    return null;
  }
  return stripTime_(parsed);
}

function normalizeDateTimeInput_(value) {
  if (!value && value !== 0) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return new Date(value.getTime());
  }
  const parsed = new Date(value);
  return isNaN(parsed) ? null : parsed;
}

function normalizeTimeInput_(value) {
  if (!value && value !== 0) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, getScriptTimeZone_(), 'HH:mm');
  }

  if (typeof value === 'number' && isFinite(value)) {
    const minutesInDay = 24 * 60;
    const totalMinutes = Math.round(Number(value) * minutesInDay);
    const normalizedMinutes = ((totalMinutes % minutesInDay) + minutesInDay) % minutesInDay;
    const hours = Math.floor(normalizedMinutes / 60);
    const minutes = normalizedMinutes % 60;
    return String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
  }

  const text = String(value).trim();
  if (!text) return '';

  let match = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (match) {
    const hours = Number(match[1]);
    const minutes = Number(match[2]);
    if (hours >= 0 && hours < 24 && minutes >= 0 && minutes < 60) {
      return String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
    }
  }

  match = text.match(/^(\d{1,2}):(\d{2})\s*([AP]M)$/i);
  if (match) {
    let hours = Number(match[1]);
    const minutes = Number(match[2]);
    const meridiem = String(match[3] || '').toUpperCase();
    if (hours >= 1 && hours <= 12 && minutes >= 0 && minutes < 60) {
      if (meridiem === 'PM' && hours !== 12) hours += 12;
      if (meridiem === 'AM' && hours === 12) hours = 0;
      return String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
    }
  }

  return text;
}

function stripTime_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatDateKey_(date) {
  return Utilities.formatDate(stripTime_(date), getScriptTimeZone_(), 'yyyy-MM-dd');
}

function formatDateForGuest_(date) {
  return Utilities.formatDate(stripTime_(normalizeDateInput_(date)), getScriptTimeZone_(), 'd MMM yyyy');
}

function getBedSetupOptions_() {
  return BED_SETUP_VALUES.map(function(value) {
    return { value: value, label: value };
  });
}

function normalizeBedSetup_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const lower = raw.toLowerCase();
  if (lower === 'best available' || lower === 'best-available' || lower === 'best_available') return 'Best available';
  if (lower === 'double' || lower === 'dbl') return 'Double';
  if (lower === 'twin') return 'Twin';
  if (
    lower === 'single' ||
    lower === 'queen' ||
    lower === 'queen bed' ||
    lower === 'single queen' ||
    lower === 'single/queen' ||
    lower === 'single / queen' ||
    lower === 'single queen bed' ||
    lower === 'single / queen bed'
  ) return 'Single / Queen';
  if (
    lower === 'triple' ||
    lower === 'trpl' ||
    lower === 'triple / family' ||
    lower === 'triple/family' ||
    lower === 'triple-family' ||
    lower === 'family' ||
    lower === 'family room'
  ) return 'Triple';
  return raw;
}

function parseAllowedSetups_(value) {
  const raw = Array.isArray(value) ? value.join('|') : String(value || '').trim();
  if (!raw) return [];
  const seen = {};
  return raw.split(/[|,]/).map(function(part) {
    return normalizeBedSetup_(part);
  }).filter(function(item) {
    if (!item || item === 'Best available' || seen[item]) return false;
    seen[item] = true;
    return true;
  });
}

function normalizeWebsiteBedSetupPreference_(guests, value) {
  const guestCount = Math.max(1, Number(guests || 1));
  const normalized = normalizeBedSetup_(value || 'Best available') || 'Best available';
  if (guestCount <= 2) return normalized;
  if (guestCount === 3 && normalized === 'Triple') return 'Triple';
  return 'Best available';
}

function parseSellableProductRoomCodes_(value) {
  const raw = Array.isArray(value) ? value.join('|') : String(value || '').trim();
  if (!raw) return [];
  const seen = {};
  return raw.split(/[|,]/).map(function(part) {
    return canonicalizeRoomLookupKey_(part);
  }).filter(function(item) {
    if (!item || seen[item]) return false;
    seen[item] = true;
    return true;
  });
}

function getSellableProductRequiredBedSetup_(product) {
  return normalizeBedSetup_(product && (product.required_bed_setup || product.requiredBedSetup || ''));
}

function getSellableProductSupportedBedSetups_(product) {
  const requiredSetup = getSellableProductRequiredBedSetup_(product);
  const supported = parseAllowedSetups_(product && (product.supported_bed_setups || product.supportedBedSetups || requiredSetup || ''));
  return supported.length ? supported : (requiredSetup ? [requiredSetup] : []);
}

function sellableProductSupportsRequestedBedSetup_(product, requestedBedSetup) {
  const normalizedRequested = normalizeBedSetup_(requestedBedSetup || '');
  if (!normalizedRequested || normalizedRequested === 'Best available') return true;
  const supported = getSellableProductSupportedBedSetups_(product);
  if (!supported.length) return true;
  return supported.indexOf(normalizedRequested) !== -1;
}

function getSellableProductCandidateRoomIds_(product, roomIndex) {
  if (!product) return [];
  const targetRoomTypeId = resolveRoomTypeId_(String(product.room_type_id || '').trim());
  const effectiveRoomIndex = roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const codes = parseSellableProductRoomCodes_(product.candidate_room_codes || product.candidateRoomCodes || '');
  if (!codes.length) return [];
  const seen = {};
  const roomIds = [];
  codes.forEach(function(code) {
    const matchedRoom = resolveAssignedRoomFromIndex_(effectiveRoomIndex, code, targetRoomTypeId);
    if (!matchedRoom || !matchedRoom.roomId || seen[matchedRoom.roomId]) return;
    seen[matchedRoom.roomId] = true;
    roomIds.push(matchedRoom.roomId);
  });
  return roomIds;
}

function getSellableProductConstraint_(product, requestedBedSetup, roomIndex) {
  const requiredSetup = getSellableProductRequiredBedSetup_(product);
  return {
    appliedBedSetup: requiredSetup || (normalizeBedSetup_(requestedBedSetup || '') || 'Best available'),
    candidateRoomIds: getSellableProductCandidateRoomIds_(product, roomIndex)
  };
}

function getSellableRoomProductById_(productId) {
  const targetId = String(productId || '').trim();
  if (!targetId) return null;
  const match = SELLABLE_ROOM_PRODUCTS.find(function(row) {
    return String(row.product_id || '').trim() === targetId;
  });
  return match || null;
}

function getSellableRoomProductsForGuests_(guests) {
  const guestCount = Math.max(1, Number(guests || 1));
  return SELLABLE_ROOM_PRODUCTS
    .filter(function(product) {
      return guestCount >= Number(product.min_guests || 1) && guestCount <= Number(product.max_guests || 1);
    })
    .slice()
    .sort(function(a, b) {
      return Number(a.sort_order || 0) - Number(b.sort_order || 0);
    });
}

function validateSellableProductSelection_(productId, roomTypeId, guests) {
  const selectedProductId = String(productId || '').trim();
  if (!selectedProductId) return null;
  const product = getSellableRoomProductById_(selectedProductId);
  if (!product) {
    throw new Error('Selected room product is invalid.');
  }
  if (String(product.room_type_id || '').trim() !== resolveRoomTypeId_(roomTypeId)) {
    throw new Error('Selected room product does not match the chosen room type.');
  }
  const guestCount = Math.max(1, Number(guests || 1));
  if (guestCount < Number(product.min_guests || 1) || guestCount > Number(product.max_guests || 1)) {
    throw new Error('Selected room product is not valid for this guest count.');
  }
  return product;
}

function getRoomPhysicalMaxGuests_(room) {
  if (!room) return 0;
  if (String(room.roomTypeId || '').trim() === 'COTTAGE') return 4;
  const allowedSetups = Array.isArray(room.allowedSetups) ? room.allowedSetups : parseAllowedSetups_(room.allowedSetups || '');
  if (allowedSetups.indexOf('Triple') !== -1) return 3;
  return 2;
}

function isRoomOccupancyCompatibleWithGuests_(room, guests) {
  return getRoomPhysicalMaxGuests_(room) >= Math.max(1, Number(guests || 1));
}

function isRequestedSetupCompatibleWithRoomForGuests_(requestedSetup, guests, room, options) {
  const config = options || {};
  const normalizedSetup = normalizeBedSetup_(requestedSetup || '');
  const guestCount = Math.max(1, Number(guests || 1));
  if (!room) return false;
  if (!normalizedSetup || normalizedSetup === 'Best available') return true;

  if (guestCount > 2 && normalizedSetup !== 'Triple' && !toBoolean_(config.strictLargeGroupSetup)) {
    return true;
  }

  if (String(room.roomTypeId || '').trim() === 'COTTAGE') {
    return normalizedSetup === 'Double' || normalizedSetup === 'Twin';
  }

  const allowedSetups = Array.isArray(room.allowedSetups) ? room.allowedSetups : parseAllowedSetups_(room.allowedSetups || '');
  if (!allowedSetups.length) return true;
  return allowedSetups.indexOf(normalizedSetup) !== -1;
}

function roomMatchesStayRequirements_(room, guests, requestedSetup, options) {
  return isRoomOccupancyCompatibleWithGuests_(room, guests) &&
    isRequestedSetupCompatibleWithRoomForGuests_(requestedSetup, guests, room, options);
}

function getActivePhysicalRoomsForType_(roomTypeId, roomIndex) {
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const effectiveRoomIndex = roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  return (effectiveRoomIndex.byType[targetRoomTypeId] || []).filter(function(room) {
    return room && room.active;
  });
}

function buildStayDemandUnitsForDate_(date, roomTypeId, options) {
  const config = options || {};
  const targetDate = normalizeDateInput_(date);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const excludeBookingId = String(config.excludeBookingId || '').trim();
  const rows = config.bookingRows || getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const effectiveRoomIndex = config.roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const units = [];

  (rows || []).forEach(function(row) {
    const bookingId = String(row.booking_id || '').trim();
    if (excludeBookingId && bookingId === excludeBookingId) return;
    if (String(row.room_type_id || '').trim() !== targetRoomTypeId) return;

    const status = normalizeBookingStatus_(row.status);
    if (BOOKING_STATUSES_COUNTED.indexOf(status) === -1) return;

    const checkIn = normalizeDateInput_(row.check_in);
    const checkOut = getInventoryEffectiveCheckOut_(row);
    if (!checkIn || !checkOut) return;
    if (!staysOverlap_(targetDate, addDays_(targetDate, 1), checkIn, checkOut)) return;

    const qtyRooms = Math.max(1, Number(row.qty_rooms || 1));
    const totalGuests = Math.max(1, Number(row.guests || Number(row.adults || 0) + Number(row.children || 0) || 1));
    const guests = getGuestsPerRoom_(totalGuests, qtyRooms);
    const bedSetup = normalizeBedSetup_(row.bed_setup || '');
    const assignedRoom = qtyRooms === 1
      ? resolveAssignedRoomFromIndex_(effectiveRoomIndex, row.room_identifier, targetRoomTypeId)
      : null;

    for (let index = 0; index < qtyRooms; index++) {
      units.push({
        unitId: bookingId + '#' + index,
        bookingId: bookingId,
        guests: guests,
        bedSetup: bedSetup,
        fixedRoomId: assignedRoom && index === 0 ? assignedRoom.roomId : ''
      });
    }
  });

  return units;
}

function buildRequestedStayUnits_(prefix, qtyRooms, guests, bedSetup) {
  const count = Math.max(0, Number(qtyRooms || 0));
  const rows = [];
  for (let index = 0; index < count; index++) {
    rows.push({
      unitId: String(prefix || 'requested') + '#' + index,
      bookingId: '',
      guests: Math.max(1, Number(guests || 1)),
      bedSetup: normalizeBedSetup_(bedSetup || ''),
      fixedRoomId: ''
    });
  }
  return rows;
}

function buildDemandUnitCandidateRoomIds_(unit, rooms, options) {
  const config = options || {};
  if (!unit || !Array.isArray(rooms) || !rooms.length) return [];
  const restrictedRoomIds = Array.isArray(config.candidateRoomIds)
    ? config.candidateRoomIds.map(function(value) {
        return String(value || '').trim();
      }).filter(Boolean)
    : [];
  const compatibleRoomIds = rooms.filter(function(room) {
    return roomMatchesStayRequirements_(room, unit.guests, unit.bedSetup, config);
  }).map(function(room) {
    return room.roomId;
  }).filter(function(roomId) {
    return !restrictedRoomIds.length || restrictedRoomIds.indexOf(roomId) !== -1;
  });
  if (unit.fixedRoomId) {
    const fixedRoom = rooms.find(function(room) {
      return String(room.roomId || '').trim() === String(unit.fixedRoomId || '').trim();
    });
    if (fixedRoom && roomMatchesStayRequirements_(fixedRoom, unit.guests, unit.bedSetup, config)) {
      return [fixedRoom.roomId];
    }
    // Legacy assignments can be stale or operationally invalid. Keep availability
    // checks resilient by falling back to any compatible room instead of marking
    // the whole date infeasible.
    if (toBoolean_(config.allowFixedFallbackToCompatiblePool)) {
      return compatibleRoomIds;
    }
    return [];
  }
  return compatibleRoomIds;
}

function canAllocateDemandUnitsToRooms_(rooms, units, blockedCount) {
  const effectiveRooms = Array.isArray(rooms) ? rooms : [];
  const effectiveUnits = Array.isArray(units) ? units : [];
  const unavailableRooms = Math.max(0, Number(blockedCount || 0));
  if (effectiveUnits.length > Math.max(0, effectiveRooms.length - unavailableRooms)) {
    return false;
  }

  const usedRoomIds = {};
  const flexibleUnits = [];

  for (let index = 0; index < effectiveUnits.length; index++) {
    const unit = effectiveUnits[index];
    const candidateIds = Array.isArray(unit.candidateRoomIds) ? unit.candidateRoomIds.slice() : [];
    if (!candidateIds.length) return false;
    if (unit.fixedRoomId) {
      const fixedRoomId = String(unit.fixedRoomId || '').trim();
      if (!fixedRoomId || candidateIds.indexOf(fixedRoomId) === -1) return false;
      if (usedRoomIds[fixedRoomId]) return false;
      usedRoomIds[fixedRoomId] = true;
      continue;
    }
    flexibleUnits.push({
      unitId: unit.unitId,
      candidateRoomIds: candidateIds
    });
  }

  flexibleUnits.sort(function(a, b) {
    return a.candidateRoomIds.length - b.candidateRoomIds.length;
  });

  function assignNext(position) {
    if (position >= flexibleUnits.length) return true;
    const unit = flexibleUnits[position];
    for (let idx = 0; idx < unit.candidateRoomIds.length; idx++) {
      const roomId = unit.candidateRoomIds[idx];
      if (usedRoomIds[roomId]) continue;
      usedRoomIds[roomId] = true;
      if (assignNext(position + 1)) return true;
      delete usedRoomIds[roomId];
    }
    return false;
  }

  return assignNext(0);
}

function getCompatibleAvailabilityForDate_(date, roomTypeId, guests, requestedSetup, options) {
  const config = options || {};
  const targetDate = normalizeDateInput_(date);
  const targetRoomTypeId = resolveRoomTypeId_(roomTypeId);
  const effectiveRoomIndex = config.roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const rooms = getActivePhysicalRoomsForType_(targetRoomTypeId, effectiveRoomIndex);
  const unitCacheKey = getDateRoomKey_(targetDate, targetRoomTypeId);
  const cachedUnits = config.demandUnitsByDateRoom &&
    Object.prototype.hasOwnProperty.call(config.demandUnitsByDateRoom, unitCacheKey)
    ? config.demandUnitsByDateRoom[unitCacheKey]
    : null;
  const excludeBookingId = String(config.excludeBookingId || '').trim();
  const demandUnits = cachedUnits
    ? cachedUnits.filter(function(unit) {
        return !excludeBookingId || String(unit.bookingId || '').trim() !== excludeBookingId;
      })
    : buildStayDemandUnitsForDate_(targetDate, targetRoomTypeId, {
        bookingRows: config.bookingRows,
        roomIndex: effectiveRoomIndex,
        excludeBookingId: excludeBookingId
      });
  const existingUnits = demandUnits.map(function(unit) {
    const candidateRoomIds = buildDemandUnitCandidateRoomIds_(unit, rooms, {
      strictLargeGroupSetup: false,
      allowFixedFallbackToCompatiblePool: true
    });
    return Object.assign({}, unit, {
      candidateRoomIds: candidateRoomIds,
      fixedRoomId: unit.fixedRoomId && candidateRoomIds.length === 1 && candidateRoomIds[0] === unit.fixedRoomId
        ? unit.fixedRoomId
        : ''
    });
  });
  const blockedCount = getBlockedDatesCount_(targetDate, targetRoomTypeId, config.blockedCountMap);
  const requestedUnits = buildRequestedStayUnits_('requested', 1, guests, requestedSetup);
  const requestedCandidateRoomIds = requestedUnits.length
    ? buildDemandUnitCandidateRoomIds_(requestedUnits[0], rooms, {
        strictLargeGroupSetup: false,
        candidateRoomIds: config.requestedCandidateRoomIds
      })
    : [];
  const compatibleInventory = requestedCandidateRoomIds.length;
  const existingFeasible = canAllocateDemandUnitsToRooms_(rooms, existingUnits, blockedCount);
  const maxPhysicalCapacity = Math.max(0, rooms.length - blockedCount - existingUnits.length);
  let maxAdditionalRooms = 0;

  if (existingFeasible && compatibleInventory > 0 && maxPhysicalCapacity > 0) {
    for (let qty = 1; qty <= maxPhysicalCapacity; qty++) {
      const candidateUnits = existingUnits.concat(buildRequestedStayUnits_('requested', qty, guests, requestedSetup).map(function(unit) {
        return Object.assign({}, unit, {
          candidateRoomIds: requestedCandidateRoomIds.slice()
        });
      }));
      if (!canAllocateDemandUnitsToRooms_(rooms, candidateUnits, blockedCount)) {
        break;
      }
      maxAdditionalRooms = qty;
    }
  }

  return {
    roomTypeId: targetRoomTypeId,
    date: formatDateKey_(targetDate),
    physicalInventory: rooms.length,
    compatibleInventory: compatibleInventory,
    blockedRooms: blockedCount,
    soldRooms: existingUnits.length,
    availableRooms: maxAdditionalRooms,
    existingFeasible: existingFeasible,
    requestedCandidateRoomIds: requestedCandidateRoomIds
  };
}

function groupRowsByRoomTypeId_(rows, keyName) {
  return (rows || []).reduce(function(map, row) {
    const roomTypeId = String(row && row[keyName || 'roomTypeId'] || '').trim();
    if (!roomTypeId) return map;
    if (!map[roomTypeId]) map[roomTypeId] = [];
    map[roomTypeId].push(row);
    return map;
  }, {});
}

function buildRoomTypeNameMap_(roomTypeRows) {
  return (roomTypeRows || []).reduce(function(map, row) {
    const roomTypeId = String(row.room_type_id || row.roomTypeId || '').trim();
    if (roomTypeId) {
      map[roomTypeId] = String(row.room_type_name || row.roomTypeName || roomTypeId).trim();
    }
    return map;
  }, {});
}

function buildInventoryMapFromRoomRows_(roomRows) {
  return (roomRows || []).reduce(function(map, room) {
    if (!room || !room.active || !room.roomTypeId) return map;
    map[room.roomTypeId] = Number(map[room.roomTypeId] || 0) + 1;
    return map;
  }, {});
}

function buildPublicAvailabilitySearchContext_(checkIn, checkOut, roomTypeIds, options) {
  const config = options || {};
  const timingPrefix = Object.prototype.hasOwnProperty.call(config, 'timingPrefix') ? config.timingPrefix : 'publicSearch';
  const readSheetsStartedAt = Date.now();
  const spreadsheet = getSpreadsheet_();
  const roomTypeRows = config.roomTypeRows || getSheetObjects_(SHEET_NAMES.ROOM_TYPES);
  const roomTypeNameMap = buildRoomTypeNameMap_(roomTypeRows);
  const roomRows = config.roomRows || getRoomsMasterRows_({ activeOnly: false });
  const roomIndex = config.roomIndex || buildRoomMasterIndex_(roomRows);
  const bookingRows = config.bookingRows || getBookingRowsForAvailabilityWindow_(
    getSheetOrThrow_(spreadsheet, SHEET_NAMES.BOOKINGS),
    checkIn,
    checkOut,
    roomTypeIds
  );
  const blockedRows = config.blockedRows || getBlockedRowsForAvailabilityWindow_(spreadsheet, checkIn, checkOut, roomTypeIds);
  const baseRateRowsRaw = config.baseRateRowsRaw || getSheetObjects_(SHEET_NAMES.BASE_RATES);
  const commercialControlRowsRaw = config.commercialControlRowsRaw || getSheetObjects_(SHEET_NAMES.COMMERCIAL_CONTROLS);
  const legacyRateRowsRaw = config.legacyRateRowsRaw || getSheetObjects_(SHEET_NAMES.RATES);
  const bookingCurrency = config.bookingCurrency || getDefaultBookingCurrency_();
  if (timingPrefix === 'publicSearch') {
    logTiming_('publicSearch:readSheets', readSheetsStartedAt);
  } else if (timingPrefix) {
    logTiming_(timingPrefix + ':readSheets', readSheetsStartedAt);
  }

  const buildContextStartedAt = Date.now();
  const mapBounds = {
    startDate: checkIn,
    endDateExclusive: checkOut
  };
  const baseRateRows = baseRateRowsRaw
    .map(function(row, index) {
      return normalizeBaseRateRow_(row, index + 2, roomTypeNameMap);
    })
    .filter(function(row) {
      return row.roomTypeId && row.active;
    });
  const commercialControls = commercialControlRowsRaw
    .map(function(row, index) {
      return normalizeCommercialControlRow_(row, index + 2, roomTypeNameMap);
    })
    .filter(function(row) {
      return row.controlId && row.roomTypeId && row.startDate && row.endDate && row.active;
    });
  const legacyRateRows = legacyRateRowsRaw.filter(function(row) {
    return String(row.status || '').trim().toLowerCase() === 'open';
  });
  const shouldBuildDemandUnitsByDateRoom = config.buildDemandUnitsByDateRoom === true &&
    Array.isArray(roomTypeIds) &&
    roomTypeIds.length > 0;
  const demandUnitsByDateRoom = shouldBuildDemandUnitsByDateRoom
    ? buildDemandUnitMapForAvailabilityWindow_(bookingRows, checkIn, checkOut, roomTypeIds, roomIndex)
    : null;

  const context = {
    roomTypeNameMap: roomTypeNameMap,
    roomIndex: roomIndex,
    inventoryMap: buildInventoryMapFromRoomRows_(roomRows),
    bookingRows: bookingRows,
    blockedRows: blockedRows,
    demandUnitsByDateRoom: demandUnitsByDateRoom,
    bookingCountMap: buildConfirmedBookingRoomCountMap_(bookingRows, mapBounds).byDateRoom,
    blockedCountMap: buildBlockedDateQtyMap_(blockedRows, mapBounds).byDateRoom,
    baseRateRowsByRoomType: groupRowsByRoomTypeId_(baseRateRows, 'roomTypeId'),
    commercialControlsByRoomType: groupRowsByRoomTypeId_(commercialControls, 'roomTypeId'),
    legacyRateRowsByRoomType: groupRowsByRoomTypeId_(legacyRateRows, 'room_type_id'),
    bookingCurrency: bookingCurrency
  };
  if (timingPrefix === 'publicSearch') {
    logTiming_('publicSearch:buildContext', buildContextStartedAt);
  } else if (timingPrefix) {
    logTiming_(timingPrefix + ':buildContext', buildContextStartedAt);
  }
  return context;
}

function searchSellableAvailabilityProducts_(checkIn, checkOut, guests, bedSetup) {
  const effectiveBedSetup = normalizeWebsiteBedSetupPreference_(guests, bedSetup);
  const products = getSellableRoomProductsForGuests_(guests).filter(function(product) {
    return sellableProductSupportsRequestedBedSetup_(product, effectiveBedSetup);
  });
  if (!products.length) {
    logTiming_('publicSearch:evaluateProducts:start');
    logTiming_('publicSearch:evaluateProducts:inventory', Date.now());
    logTiming_('publicSearch:evaluateProducts:pricing', Date.now());
    logTiming_('publicSearch:evaluateProducts:end', Date.now());
    return [];
  }
  const roomTypeIds = products.reduce(function(ids, product) {
    const roomTypeId = String(product && product.room_type_id || '').trim();
    if (roomTypeId && ids.indexOf(roomTypeId) === -1) ids.push(roomTypeId);
    return ids;
  }, []);
  const searchContext = buildPublicAvailabilitySearchContext_(checkIn, checkOut, roomTypeIds, {
    buildDemandUnitsByDateRoom: true
  });
  const bookingCurrency = searchContext.bookingCurrency || getDefaultBookingCurrency_();
  const sortOrderByProductId = {};
  const evaluateProductsStartedAt = Date.now();
  const timingAccumulator = {
    inventoryMs: 0,
    pricingMs: 0
  };
  logTiming_('publicSearch:evaluateProducts:start');
  const results = products.map(function(product) {
    sortOrderByProductId[product.product_id] = Number(product.sort_order || 0);
    const productRoomTypeId = String(product.room_type_id || '').trim();
    const productConstraint = getSellableProductConstraint_(product, effectiveBedSetup, searchContext.roomIndex);
    const snapshot = buildStayAvailabilityPricingSnapshot_(checkIn, checkOut, product.room_type_id, guests, {
      qtyRooms: 1,
      bedSetup: productConstraint.appliedBedSetup,
      sellableProductId: product.product_id,
      roomIndex: searchContext.roomIndex,
      candidateRoomIds: productConstraint.candidateRoomIds,
      bookingRows: searchContext.bookingRows,
      blockedRows: searchContext.blockedRows,
      bookingCountMap: searchContext.bookingCountMap,
      blockedCountMap: searchContext.blockedCountMap,
      roomTypeNameMap: searchContext.roomTypeNameMap,
      inventoryTotal: searchContext.inventoryMap[productRoomTypeId],
      physicalInventory: searchContext.inventoryMap[productRoomTypeId],
      commercialControls: searchContext.commercialControlsByRoomType[productRoomTypeId] || [],
      baseRateRows: searchContext.baseRateRowsByRoomType[productRoomTypeId] || [],
      legacyRateRows: searchContext.legacyRateRowsByRoomType[productRoomTypeId] || [],
      demandUnitsByDateRoom: searchContext.demandUnitsByDateRoom,
      currency: bookingCurrency,
      timingAccumulator: timingAccumulator
    });
    const nights = Math.max(0, Number(snapshot.breakdown ? snapshot.breakdown.length : 0));
    const availabilityState = buildPublicProductAvailabilityState_(snapshot);
    return {
      productId: product.product_id,
      roomTypeId: product.room_type_id,
      roomTypeName: snapshot.roomTypeName,
      productName: product.product_name,
      productLabel: product.product_label,
      standardLabel: product.standard_label || '',
      productSummary: product.product_summary || '',
      bedSummary: product.bed_summary || '',
      imageSrc: product.image_src || '',
      imageAlt: product.image_alt || '',
      highlights: Array.isArray(product.highlights) ? product.highlights.slice(0, 3) : [],
      capacityLabel: product.capacity_label,
      occupancyBucket: Number(product.occupancy_bucket || guests || 1),
      guests: Math.max(1, Number(guests || 1)),
      bedSetup: snapshot.bedSetup || productConstraint.appliedBedSetup || 'Best available',
      availableRooms: Math.max(0, Number(snapshot.availableRooms || 0)),
      available: availabilityState.bookable,
      bookable: availabilityState.bookable,
      disabledReason: availabilityState.reason,
      buttonLabel: availabilityState.buttonLabel,
      availabilityLabel: availabilityState.label,
      availabilityTone: availabilityState.tone,
      estimatedPrice: snapshot.estimatedPrice,
      directPrice: snapshot.estimatedPrice,
      comparisonPrice: snapshot.directBookingOffer && snapshot.directBookingOffer.hasSaving ? snapshot.directBookingOffer.comparisonPrice : null,
      publicReferencePrice: snapshot.directBookingOffer && snapshot.directBookingOffer.hasSaving ? snapshot.directBookingOffer.publicReferencePrice : null,
      directDiscountType: snapshot.directBookingOffer ? snapshot.directBookingOffer.directDiscountType : 'None',
      directDiscountValue: snapshot.directBookingOffer ? snapshot.directBookingOffer.directDiscountValue : '',
      savingsAmount: snapshot.directBookingOffer && snapshot.directBookingOffer.hasSaving ? snapshot.directBookingOffer.savingsAmount : 0,
      savingsPercentage: snapshot.directBookingOffer && snapshot.directBookingOffer.hasSaving ? snapshot.directBookingOffer.savingsPercentage : 0,
      savingsLabel: snapshot.directBookingOffer && snapshot.directBookingOffer.hasSaving ? snapshot.directBookingOffer.savingsLabel : '',
      nightlyPrice: snapshot.estimatedPrice != null && nights > 0 ? roundCurrency_(snapshot.estimatedPrice / nights) : null,
      nights: nights,
      currency: bookingCurrency,
      pricingSource: snapshot.pricingSource,
      pricingReferenceId: snapshot.pricingReferenceId,
      pricingNotes: snapshot.pricingNotes || []
    };
  }).filter(Boolean).sort(function(a, b) {
    if (!!a.bookable !== !!b.bookable) return a.bookable ? -1 : 1;
    const orderDelta = Number(sortOrderByProductId[a.productId] || 0) - Number(sortOrderByProductId[b.productId] || 0);
    if (orderDelta !== 0) return orderDelta;
    return Number(a.estimatedPrice || 0) - Number(b.estimatedPrice || 0);
  });
  logTiming_('publicSearch:evaluateProducts:inventory', Date.now() - Number(timingAccumulator.inventoryMs || 0));
  logTiming_('publicSearch:evaluateProducts:pricing', Date.now() - Number(timingAccumulator.pricingMs || 0));
  logTiming_('publicSearch:evaluateProducts:end', evaluateProductsStartedAt);
  return results;
}

function isRoomMasterActive_(roomRow) {
  const raw = String(roomRow && roomRow.active != null ? roomRow.active : 'Yes').trim().toLowerCase();
  return raw === '' || raw === 'yes' || raw === 'true' || raw === '1' || raw === 'active';
}

function getCanonicalRoomCompatibilityCode_(room) {
  const roomTypeId = String(room && room.roomTypeId || '').trim();
  const rawCode = String(room && room.roomCode || '').trim().toUpperCase();
  if (roomTypeId === 'PEAKBAL' && Object.prototype.hasOwnProperty.call(ROOM_MASTER_COMPATIBILITY_BY_TYPE_AND_CODE.PEAKBAL, rawCode)) {
    return rawCode;
  }
  if (roomTypeId === 'PEAKMV') {
    const directPeakCode = rawCode.match(/^N?([15678])$/);
    if (directPeakCode) return directPeakCode[1];
    const roomAliases = [
      room.roomId,
      room.roomName
    ];
    for (let index = 0; index < roomAliases.length; index++) {
      const aliasKey = canonicalizeRoomLookupKey_(roomAliases[index]);
      const aliasMatch = aliasKey.match(/(?:^|PEAK)N?([15678])$/);
      if (aliasMatch) return aliasMatch[1];
    }
  }
  return rawCode;
}

function applyRoomMasterCompatibilityOverride_(room) {
  if (!room || !room.roomTypeId) return room;
  const byCode = ROOM_MASTER_COMPATIBILITY_BY_TYPE_AND_CODE[room.roomTypeId];
  if (!byCode) return room;
  const compatibilityCode = getCanonicalRoomCompatibilityCode_(room);
  const override = byCode[compatibilityCode];
  if (!override) return room;
  return Object.assign({}, room, {
    roomCode: override.roomCode,
    roomName: override.roomName,
    defaultSetup: normalizeBedSetup_(override.defaultSetup),
    allowedSetups: parseAllowedSetups_(override.allowedSetups)
  });
}

function normalizeRoomMasterRow_(row) {
  const roomTypeId = String(row.room_type_id || row.roomTypeId || '').trim();
  const normalized = {
    roomId: String(row.room_id || row.roomId || '').trim(),
    roomCode: String(row.room_code || row.roomCode || '').trim(),
    roomName: String(row.room_name || row.roomName || row.room_code || row.roomCode || '').trim(),
    roomTypeId: roomTypeId,
    roomTypeName: String(row.room_type_name || row.roomTypeName || row.room_type_id || row.roomTypeId || '').trim(),
    defaultSetup: normalizeBedSetup_(row.default_setup || row.defaultSetup || ''),
    allowedSetups: parseAllowedSetups_(row.allowed_setups || row.allowedSetups || ''),
    active: isRoomMasterActive_(row),
    sortOrder: Number(row.sort_order || row.sortOrder || 0),
    notes: String(row.notes || '').trim()
  };
  return applyRoomMasterCompatibilityOverride_(normalized);
}

function canonicalizeRoomLookupKey_(value) {
  return String(value || '').trim().toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function buildRoomLookupKeys_(room) {
  const aliases = [
    room.roomId,
    room.roomCode,
    room.roomName,
    room.roomTypeName && room.roomCode ? (room.roomTypeName + ' ' + room.roomCode) : '',
    room.roomTypeId && room.roomCode ? (room.roomTypeId + ' ' + room.roomCode) : ''
  ];
  const byTypeAlias = room.roomTypeId === 'PEAKBAL'
    ? ['Peak Balcony ' + room.roomCode, 'Peak with Balcony ' + room.roomCode]
    : room.roomTypeId === 'PEAKMV'
      ? ['Peak ' + room.roomCode].concat(String(room.roomCode || '').match(/^[0-9]+$/)
          ? ['N' + room.roomCode, 'Peak N' + room.roomCode]
          : [])
      : room.roomTypeId === 'CLASSIC'
        ? ['Classic ' + room.roomCode]
        : room.roomTypeId === 'COTTAGE'
          ? ['Cottage ' + room.roomCode, 'Peak Cottage ' + room.roomCode]
          : [];
  return aliases.concat(byTypeAlias).map(canonicalizeRoomLookupKey_).filter(Boolean);
}

function getRoomsMasterRows_(options) {
  const config = options || {};
  const spreadsheet = getSpreadsheet_();
  const roomSheet = spreadsheet.getSheetByName(SHEET_NAMES.ROOMS);
  if (!roomSheet || roomSheet.getLastRow() < 2) {
    return [];
  }
  const rows = getSheetObjects_(SHEET_NAMES.ROOMS).map(normalizeRoomMasterRow_);
  return rows
    .filter(function(room) {
      return room.roomId && (!config.activeOnly || room.active);
    })
    .sort(function(a, b) {
      const sortDelta = Number(a.sortOrder || 0) - Number(b.sortOrder || 0);
      if (sortDelta !== 0) return sortDelta;
      return String(a.roomName || a.roomCode || '').localeCompare(String(b.roomName || b.roomCode || ''));
    });
}

function buildRoomMasterIndex_(roomRows) {
  const rows = (roomRows || []).map(normalizeRoomMasterRow_);
  const index = {
    rows: rows,
    byLookup: {},
    byType: {}
  };

  rows.forEach(function(room) {
    if (!room.roomId) return;
    if (!index.byType[room.roomTypeId]) index.byType[room.roomTypeId] = [];
    index.byType[room.roomTypeId].push(room);
    buildRoomLookupKeys_(room).forEach(function(key) {
      if (key && !index.byLookup[key]) {
        index.byLookup[key] = room;
      }
    });
  });

  return index;
}

function buildRoomMasterBootstrapRows_() {
  return buildRoomMasterBootstrapRowsFromRows_(getRoomsMasterRows_({ activeOnly: true }));
}

function buildRoomMasterBootstrapRowsFromRows_(roomRows) {
  return (roomRows || []).filter(function(room) {
    return room && room.active;
  }).map(function(room) {
    return {
      roomId: room.roomId,
      roomCode: room.roomCode,
      roomName: room.roomName,
      roomTypeId: room.roomTypeId,
      roomTypeName: room.roomTypeName,
      defaultSetup: room.defaultSetup,
      allowedSetups: room.allowedSetups,
      active: room.active,
      sortOrder: room.sortOrder,
      notes: room.notes
    };
  });
}

function getRoomInventoryMap_() {
  return getRoomsMasterRows_({ activeOnly: true }).reduce(function(map, room) {
    map[room.roomTypeId] = Number(map[room.roomTypeId] || 0) + 1;
    return map;
  }, {});
}

function getActiveRoomTypeCatalog_() {
  return buildActiveRoomTypeCatalogFromRows_(getSheetObjects_(SHEET_NAMES.ROOM_TYPES), getRoomsMasterRows_({ activeOnly: true }));
}

function buildActiveRoomTypeCatalogFromRows_(roomTypeRows, roomRows) {
  const inventoryMap = (roomRows || []).reduce(function(map, room) {
    if (!room || !room.active || !room.roomTypeId) return map;
    map[room.roomTypeId] = Number(map[room.roomTypeId] || 0) + 1;
    return map;
  }, {});
  return (roomTypeRows || [])
    .filter(isActiveRoomType_)
    .map(function(row) {
      const roomTypeId = String(row.room_type_id || '').trim();
      return {
        roomTypeId: roomTypeId,
        roomTypeName: String(row.room_type_name || row.room_type_id || '').trim(),
        inventoryTotal: Object.prototype.hasOwnProperty.call(inventoryMap, roomTypeId)
          ? Number(inventoryMap[roomTypeId] || 0)
          : Number(row.inventory_total || 0),
        maxGuests: Number(row.max_guests || 0)
      };
    })
    .filter(function(row) {
      return row.roomTypeId;
    });
}

function resolveAssignedRoomFromIndex_(roomIndex, roomIdentifier, roomTypeId) {
  const raw = String(roomIdentifier || '').trim();
  if (!raw || !roomIndex) return null;
  const lookupKey = canonicalizeRoomLookupKey_(raw);
  if (!lookupKey) return null;

  if (roomTypeId && roomIndex.byType[roomTypeId]) {
    const scoped = roomIndex.byType[roomTypeId].find(function(room) {
      return buildRoomLookupKeys_(room).indexOf(lookupKey) !== -1;
    });
    if (scoped) return scoped;
  }

  return roomIndex.byLookup[lookupKey] || null;
}

function isRequestedSetupCompatibleWithRoom_(requestedSetup, room) {
  return isRequestedSetupCompatibleWithRoomForGuests_(requestedSetup, 2, room, {
    strictLargeGroupSetup: false
  });
}

function staysOverlap_(startA, endA, startB, endB) {
  if (!startA || !endA || !startB || !endB) return false;
  return stripTime_(startA).getTime() < stripTime_(endB).getTime() &&
    stripTime_(startB).getTime() < stripTime_(endA).getTime();
}

function findRoomAssignmentConflict_(room, bookingRows, roomIndex, options) {
  const config = options || {};
  const checkIn = normalizeDateInput_(config.checkIn);
  const checkOut = normalizeDateInput_(config.checkOut);
  const excludeBookingId = String(config.excludeBookingId || '').trim();
  if (!room || !checkIn || !checkOut) return null;

  const targetLookupKey = canonicalizeRoomLookupKey_(room.roomCode || room.roomId || '');
  const rows = bookingRows || getSheetObjects_(SHEET_NAMES.BOOKINGS);
  let conflict = null;

  rows.some(function(row) {
    const bookingId = String(row.booking_id || '').trim();
    if (excludeBookingId && bookingId === excludeBookingId) return false;

    const status = normalizeBookingStatus_(row.status);
    if (status !== BOOKING_STATUS_CONFIRMED && status !== BOOKING_STATUS_IN_HOUSE) return false;

    const existingCheckIn = normalizeDateInput_(row.check_in);
    const existingCheckOut = normalizeDateInput_(row.check_out);
    if (!staysOverlap_(checkIn, checkOut, existingCheckIn, existingCheckOut)) return false;

    const assignedRoom = resolveAssignedRoomFromIndex_(roomIndex, row.room_identifier, String(row.room_type_id || '').trim());
    const sameRoom = assignedRoom
      ? assignedRoom.roomId === room.roomId
      : canonicalizeRoomLookupKey_(row.room_identifier) === targetLookupKey;
    if (!sameRoom) return false;

    conflict = {
      bookingId: bookingId,
      guestName: String(row.guest_name || '').trim(),
      status: status,
      checkIn: existingCheckIn ? formatDateKey_(existingCheckIn) : '',
      checkOut: existingCheckOut ? formatDateKey_(existingCheckOut) : '',
      roomIdentifier: String(row.room_identifier || '').trim()
    };
    return true;
  });

  return conflict;
}

function buildRoomAssignmentOptions_(params, bookings, roomIndex) {
  const config = params || {};
  const targetRoomTypeId = String(config.roomTypeId || '').trim();
  if (!targetRoomTypeId) return [];

  const effectiveRoomIndex = roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const candidateRooms = (effectiveRoomIndex.byType[targetRoomTypeId] || []).filter(function(room) {
    return room && room.active;
  });
  const currentRaw = String(config.currentRoomIdentifier || '').trim();
  const currentRoom = resolveAssignedRoomFromIndex_(effectiveRoomIndex, currentRaw, targetRoomTypeId);
  const currentLookupKey = canonicalizeRoomLookupKey_(currentRaw);
  const requestedGuests = Math.max(1, Number(config.guests || 1));
  const requestedSetup = normalizeBedSetup_(config.bedSetup || '');
  const effectiveBookings = bookings || getSheetObjects_(SHEET_NAMES.BOOKINGS);

  return candidateRooms.map(function(room) {
    const isCurrent = currentRoom
      ? currentRoom.roomId === room.roomId
      : (currentLookupKey && currentLookupKey === canonicalizeRoomLookupKey_(room.roomCode || room.roomId || ''));
    const compatible = roomMatchesStayRequirements_(room, requestedGuests, requestedSetup, {
      strictLargeGroupSetup: false
    });
    const conflict = findRoomAssignmentConflict_(room, effectiveBookings, effectiveRoomIndex, {
      checkIn: config.checkIn,
      checkOut: config.checkOut,
      excludeBookingId: config.excludeBookingId
    });
    const canAssign = compatible && !conflict;
    const statusCode = isCurrent
      ? (canAssign ? 'current' : 'current-issue')
      : (!compatible ? 'incompatible' : (conflict ? 'occupied' : 'available'));
    const statusLabel = statusCode === 'current'
      ? 'Currently assigned'
      : statusCode === 'current-issue'
        ? 'Currently assigned · review needed'
        : statusCode === 'occupied'
            ? 'Already allocated'
            : statusCode === 'incompatible'
              ? 'Setup or occupancy mismatch'
              : 'Available';

    return {
      roomId: room.roomId,
      roomCode: room.roomCode,
      roomName: room.roomName,
      roomTypeId: room.roomTypeId,
      roomTypeName: room.roomTypeName,
      roomMaxGuests: getRoomPhysicalMaxGuests_(room),
      defaultSetup: room.defaultSetup,
      allowedSetups: room.allowedSetups,
      notes: room.notes,
      isCurrent: isCurrent,
      isCompatible: compatible,
      canAssign: canAssign,
      statusCode: statusCode,
      statusLabel: statusLabel,
      conflict: conflict,
      conflictLabel: conflict
        ? (conflict.guestName || conflict.bookingId || 'Another booking') + ' · ' + [conflict.checkIn, conflict.checkOut].filter(Boolean).join(' to ')
        : '',
      recommended: canAssign && compatible
    };
  }).sort(function(a, b) {
    const rank = function(option) {
      if (option.isCurrent) return 0;
      if (option.canAssign && option.isCompatible) return 1;
      if (option.isCompatible && option.conflict) return 2;
      return 3;
    };
    const delta = rank(a) - rank(b);
    if (delta !== 0) return delta;
    return String(a.roomName || a.roomCode || '').localeCompare(String(b.roomName || b.roomCode || ''));
  });
}

function normalizeRoomAssignmentForSave_(roomIdentifier, roomTypeId, options) {
  const config = options || {};
  const raw = String(roomIdentifier || '').trim();
  const existingRaw = String(config.existingRoomIdentifier || '').trim();
  if (!raw) return '';

  if (Math.max(1, Number(config.qtyRooms || 1)) > 1) {
    if (existingRaw && canonicalizeRoomLookupKey_(existingRaw) === canonicalizeRoomLookupKey_(raw)) {
      return existingRaw;
    }
    throw new Error('Room assignment currently supports one physical room per booking. Leave room assignment blank when Rooms booked is greater than 1.');
  }

  const roomIndex = config.roomIndex || buildRoomMasterIndex_(getRoomsMasterRows_({ activeOnly: false }));
  const matchedRoom = resolveAssignedRoomFromIndex_(roomIndex, raw, roomTypeId);

  if (!matchedRoom) {
    if (existingRaw && canonicalizeRoomLookupKey_(existingRaw) === canonicalizeRoomLookupKey_(raw)) {
      return existingRaw;
    }
    throw new Error('Room assignment must match a room from the Rooms master.');
  }

  if (matchedRoom.roomTypeId !== String(roomTypeId || '').trim()) {
    throw new Error('Assigned room ' + matchedRoom.roomCode + ' belongs to ' + matchedRoom.roomTypeName + ', not ' + getRoomTypeNameById_(roomTypeId) + '.');
  }

  const requestedGuests = Math.max(1, Number(config.guests || 1));
  if (!roomMatchesStayRequirements_(matchedRoom, requestedGuests, config.bedSetup, {
    strictLargeGroupSetup: false
  })) {
    if (!isRoomOccupancyCompatibleWithGuests_(matchedRoom, requestedGuests)) {
      throw new Error('Room ' + matchedRoom.roomCode + ' cannot host ' + requestedGuests + ' guest(s).');
    }
    const allowedLabel = matchedRoom.allowedSetups && matchedRoom.allowedSetups.length
      ? matchedRoom.allowedSetups.join(' / ')
      : 'the selected setup';
    throw new Error('Room ' + matchedRoom.roomCode + ' is not compatible with ' + normalizeBedSetup_(config.bedSetup) + '. Allowed setups: ' + allowedLabel + '.');
  }

  const checkIn = normalizeDateInput_(config.checkIn);
  const checkOut = normalizeDateInput_(config.checkOut);
  if (checkIn && checkOut) {
    const bookingRows = config.bookingRows || getSheetObjects_(SHEET_NAMES.BOOKINGS);
    bookingRows.forEach(function(row) {
      const bookingId = String(row.booking_id || '').trim();
      if (config.excludeBookingId && bookingId === String(config.excludeBookingId || '').trim()) return;
      const status = normalizeBookingStatus_(row.status);
      if (status !== BOOKING_STATUS_CONFIRMED && status !== BOOKING_STATUS_IN_HOUSE) return;
      const existingCheckIn = normalizeDateInput_(row.check_in);
      const existingCheckOut = normalizeDateInput_(row.check_out);
      if (!staysOverlap_(checkIn, checkOut, existingCheckIn, existingCheckOut)) return;
      const existingMatchedRoom = resolveAssignedRoomFromIndex_(roomIndex, row.room_identifier, String(row.room_type_id || '').trim());
      const existingLookupKey = existingMatchedRoom
        ? existingMatchedRoom.roomId
        : canonicalizeRoomLookupKey_(row.room_identifier);
      if (!existingLookupKey) return;
      if (existingMatchedRoom && existingMatchedRoom.roomId === matchedRoom.roomId) {
        throw new Error('Room ' + matchedRoom.roomCode + ' is already assigned to ' + (row.guest_name || bookingId || 'another booking') + ' for overlapping dates.');
      }
      if (!existingMatchedRoom && existingLookupKey === canonicalizeRoomLookupKey_(matchedRoom.roomCode)) {
        throw new Error('Room ' + matchedRoom.roomCode + ' is already assigned to ' + (row.guest_name || bookingId || 'another booking') + ' for overlapping dates.');
      }
    });
  }

  return matchedRoom.roomCode || raw;
}

function ensureRoomTypesSeed_(spreadsheet) {
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.ROOM_TYPES, ['room_type_id', 'room_type_name', 'inventory_total', 'max_guests', 'active']);
  const headerMap = getHeaderMap_(sheet);
  DEFAULT_ROOM_TYPES_SEED.forEach(function(seedRow) {
    const rowNumber = findRowNumberByHeaderValue_(sheet, 'room_type_id', seedRow.room_type_id);
    const existing = rowNumber ? getRowObjectByNumber_(sheet, rowNumber) : null;
    const payload = {
      room_type_id: seedRow.room_type_id,
      room_type_name: seedRow.room_type_name,
      inventory_total: seedRow.inventory_total,
      max_guests: Math.max(Number(existing && existing.max_guests || 0), Number(seedRow.max_guests || 0)),
      active: seedRow.active
    };
    if (rowNumber) {
      updateObjectRow_(sheet, rowNumber, payload);
    } else {
      appendObjectRow_(sheet, headerMap, payload);
    }
  });
}

function ensureRoomsMasterSeed_(spreadsheet) {
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.ROOMS, ROOMS_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  DEFAULT_ROOMS_MASTER.forEach(function(room) {
    const rowNumber = findRowNumberByHeaderValue_(sheet, 'room_id', room.room_id);
    if (rowNumber) {
      updateObjectRow_(sheet, rowNumber, room);
    } else {
      appendObjectRow_(sheet, headerMap, room);
    }
  });
}

function ensureBaseRatesSeed_(spreadsheet) {
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BASE_RATES, BASE_RATES_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  DEFAULT_ROOM_TYPES_SEED.forEach(function(seedRow) {
    const rowNumber = findRowNumberByHeaderValue_(sheet, 'room_type_id', seedRow.room_type_id);
    const existing = rowNumber ? getRowObjectByNumber_(sheet, rowNumber) : null;
    const payload = {
      room_type_id: seedRow.room_type_id,
      room_type_name: seedRow.room_type_name,
      base_rate: existing && String(existing.base_rate || '').trim() !== ''
        ? roundCurrency_(existing.base_rate)
        : getLegacyFallbackBaseRateForRoomType_(seedRow.room_type_id),
      extra_guest_fee: existing && String(existing.extra_guest_fee == null ? '' : existing.extra_guest_fee).trim() !== ''
        ? roundCurrency_(existing.extra_guest_fee || 0)
        : (function() {
            const legacyFallback = getLegacyRateFallbackInfo_(seedRow.room_type_id, new Date());
            return legacyFallback.found ? roundCurrency_(legacyFallback.extraGuestFee || 0) : '';
          })(),
      active: existing && String(existing.active || '').trim() ? existing.active : 'Yes',
      updated_at: existing && existing.updated_at ? existing.updated_at : new Date()
    };
    if (rowNumber) {
      updateObjectRow_(sheet, rowNumber, payload);
    } else {
      appendObjectRow_(sheet, headerMap, payload);
    }
  });
}

function syncRoomTypesInventoryFromRooms_(spreadsheet) {
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.ROOM_TYPES, ['room_type_id', 'room_type_name', 'inventory_total', 'max_guests', 'active']);
  const inventoryMap = getRoomInventoryMap_();
  DEFAULT_ROOM_TYPES_SEED.forEach(function(seedRow) {
    const rowNumber = findRowNumberByHeaderValue_(sheet, 'room_type_id', seedRow.room_type_id);
    const payload = {
      room_type_id: seedRow.room_type_id,
      room_type_name: seedRow.room_type_name,
      inventory_total: Number(inventoryMap[seedRow.room_type_id] || seedRow.inventory_total || 0),
      max_guests: seedRow.max_guests,
      active: 'Yes'
    };
    if (rowNumber) {
      updateObjectRow_(sheet, rowNumber, payload);
    } else {
      const headerMap = getHeaderMap_(sheet);
      appendObjectRow_(sheet, headerMap, payload);
    }
  });
}

function ensureMiniPmsStructure_(spreadsheet) {
  const structureStartedAt = Date.now();
  logTiming_('ensureMiniPmsStructure_:start');
  try {
  const ensuredSheets = [];
  const minimumRoomTypeHeaders = ['room_type_id', 'room_type_name', 'inventory_total', 'max_guests', 'active'];
  const availabilityCacheHeaders = ['date', 'room_type_id', 'inventory_total', 'booked_confirmed', 'blocked', 'remaining', 'status'];
  const settingsHeaders = ['key', 'value'];

  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.ROOM_TYPES, minimumRoomTypeHeaders).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.ROOMS, ROOMS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BASE_RATES, BASE_RATES_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.RATES, RATES_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.COMMERCIAL_CONTROLS, COMMERCIAL_CONTROLS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.EVENT_FLAGS, EVENT_FLAGS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.COMPETITOR_TRACKER, COMPETITOR_TRACKER_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.RECOMMENDATION_ACTION_LOG, RECOMMENDATION_ACTION_LOG_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.OTA_UPDATE_WORKFLOW, OTA_UPDATE_WORKFLOW_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BOOKINGS, BOOKINGS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BLOCKED_DATES, BLOCKED_DATES_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.AVAILABILITY_CACHE, availabilityCacheHeaders).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.REQUESTS, REQUESTS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.SETTINGS, settingsHeaders).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BOOKING_NIGHTS, BOOKING_NIGHTS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.DAILY_STATS, DAILY_STATS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.OTB_SNAPSHOTS, OTB_SNAPSHOTS_HEADERS).getName());
  ensuredSheets.push(ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.WEBSITE_FEEDBACK, WEBSITE_FEEDBACK_HEADERS).getName());

  let dashboard = spreadsheet.getSheetByName(SHEET_NAMES.DASHBOARD);
  if (!dashboard) {
    dashboard = spreadsheet.insertSheet(SHEET_NAMES.DASHBOARD);
  }
  dashboard.setFrozenRows(3);
  ensuredSheets.push(dashboard.getName());

  ensureRoomTypesSeed_(spreadsheet);
  ensureRoomsMasterSeed_(spreadsheet);
  ensureBaseRatesSeed_(spreadsheet);
  syncRoomTypesInventoryFromRooms_(spreadsheet);

  return ensuredSheets;
  } finally {
    logTiming_('ensureMiniPmsStructure_:end', structureStartedAt);
  }
}

function styleHeaderRow_(sheet, rowNumber, totalColumns) {
  if (!sheet || !totalColumns) return;
  sheet.getRange(rowNumber, 1, 1, totalColumns)
    .setFontWeight('bold')
    .setBackground('#edf2ef')
    .setFontColor('#244531');
}

function parseBookingIdCounterState_(raw) {
  const text = String(raw || '').trim();
  if (!text) return { sequence: 0, lastRow: 0 };
  try {
    const parsed = JSON.parse(text);
    return {
      sequence: Math.max(0, Number(parsed.sequence || 0)),
      lastRow: Math.max(0, Number(parsed.lastRow || 0))
    };
  } catch (error) {
    return {
      sequence: Math.max(0, Number(text || 0)),
      lastRow: 0
    };
  }
}

function getMaxSequentialIdForPrefix_(sheet, headerMap, headerKey, prefix) {
  if (!Object.prototype.hasOwnProperty.call(headerMap, headerKey)) {
    throw new Error('Sheet is missing required header: ' + headerKey);
  }

  const idColumn = headerMap[headerKey] + 1;
  const lastRow = sheet.getLastRow();
  let maxSequence = 0;
  if (lastRow < 2) return maxSequence;

  const values = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();
  values.forEach(function(row) {
    const currentValue = String(row[0] || '').trim();
    if (currentValue.indexOf(prefix) !== 0) return;

    const match = currentValue.match(/-(\d+)$/);
    if (!match) return;

    maxSequence = Math.max(maxSequence, Number(match[1] || 0));
  });
  return maxSequence;
}

function generateBookingId_(sheet, headerMap, now) {
  if (!Object.prototype.hasOwnProperty.call(headerMap, 'booking_id')) {
    throw new Error('Sheet is missing required header: booking_id');
  }

  const dateKey = Utilities.formatDate(now, getScriptTimeZone_(), 'yyyyMMdd');
  const prefix = 'BKG-' + dateKey + '-';
  const propertyKey = BOOKING_ID_COUNTER_PROPERTY_PREFIX + dateKey;
  const properties = PropertiesService.getScriptProperties();
  const state = parseBookingIdCounterState_(properties.getProperty(propertyKey));
  const lastRow = sheet.getLastRow();
  let sequence = Number(state.sequence || 0);

  if (!sequence || lastRow > Number(state.lastRow || 0)) {
    sequence = Math.max(sequence, getMaxSequentialIdForPrefix_(sheet, headerMap, 'booking_id', prefix));
  }

  const nextSequence = sequence + 1;
  properties.setProperty(propertyKey, JSON.stringify({
    sequence: nextSequence,
    lastRow: lastRow + 1,
    updatedAt: new Date().toISOString()
  }));
  return prefix + padNumber_(nextSequence, 3);
}

function generateBlockedId_(sheet, headerMap, now) {
  return generateSequentialId_(sheet, headerMap, 'blocked_id', 'BLK', now);
}

function generateCompetitorTrackerEntryId_(sheet, headerMap, now) {
  return generateSequentialId_(sheet, headerMap, 'entry_id', 'CMP', now);
}

function generateEventFlagId_(sheet, headerMap, now) {
  return generateSequentialId_(sheet, headerMap, 'event_id', 'EVT', now);
}

function generateRecommendationActionLogId_(sheet, headerMap, now) {
  return generateSequentialId_(sheet, headerMap, 'log_id', 'RLOG', now);
}

function generateOtaWorkflowUpdateId_(sheet, headerMap, now) {
  return generateSequentialId_(sheet, headerMap, 'ota_update_id', 'OTA', now);
}

function getBookingSourceOptions_() {
  return BOOKING_SOURCE_VALUES.map(function(value) {
    return { value: value, label: value };
  });
}

function getPaymentStatusOptions_() {
  return PAYMENT_STATUS_VALUES.map(function(value) {
    return { value: value, label: value };
  });
}

function getPaymentMethodOptions_() {
  return PAYMENT_METHOD_VALUES.map(function(value) {
    return { value: value, label: value };
  });
}

function getCommercialRuleTypeOptions_() {
  return COMMERCIAL_RULE_TYPE_VALUES.map(function(value) {
    return {
      value: value,
      label: value === 'special' ? 'Special' : 'Seasonal'
    };
  });
}

function getReportingCurrency_() {
  return DEFAULT_REPORTING_CURRENCY;
}

function getDefaultBookingCurrency_() {
  try {
    const settings = getSettings() || {};
    return normalizeCurrencyCode_(settings.booking_currency || settings.currency || DEFAULT_CURRENCY) || DEFAULT_CURRENCY;
  } catch (error) {
    return DEFAULT_CURRENCY;
  }
}

function getBookingCurrencyOptions_() {
  return getBookingCurrencyOptionsForDefault_(getDefaultBookingCurrency_());
}

function getBookingCurrencyOptionsForDefault_(defaultCurrency) {
  const seen = {};
  return [getReportingCurrency_(), defaultCurrency].concat(BOOKING_CURRENCY_VALUES).filter(function(value) {
    const code = normalizeCurrencyCode_(value);
    if (!code || seen[code]) return false;
    seen[code] = true;
    return true;
  }).map(function(value) {
    return { value: value, label: value };
  });
}

function generateCommercialControlId_(sheet, headerMap, now) {
  return generateSequentialId_(sheet, headerMap, 'control_id', 'CTL', now);
}

function normalizeCurrencyCode_(value) {
  return String(value || '').trim().toUpperCase().replace(/[^A-Z]/g, '').slice(0, 3);
}

function roundFxRate_(value) {
  const numeric = Number(value || 0);
  if (!isFinite(numeric)) return 0;
  return Math.round(numeric * 1000000) / 1000000;
}

function getStoredBookingCurrency_(row, defaultCurrency) {
  return normalizeCurrencyCode_(row && (row.booking_currency || row.bookingCurrency || row.currency)) ||
    normalizeCurrencyCode_(defaultCurrency) ||
    getDefaultBookingCurrency_();
}

function getStoredBookingOriginalValue_(row) {
  if (!row) return 0;
  if (row.booking_value_original != null && row.booking_value_original !== '') {
    return roundCurrency_(row.booking_value_original);
  }
  return roundCurrency_(row.booking_value || 0);
}

function getOperationalBookingValueGbp_(row, options) {
  const config = options || {};
  if (!row) return 0;
  if (row.booking_value_gbp != null && row.booking_value_gbp !== '') {
    return roundCurrency_(row.booking_value_gbp);
  }
  if (row.booking_value != null && row.booking_value !== '') {
    return roundCurrency_(row.booking_value);
  }
  const originalValue = getStoredBookingOriginalValue_(row);
  const fxRate = Number(row.fx_rate_to_gbp || 0);
  if (fxRate > 0) {
    return convertAmountToGbp_(originalValue, getStoredBookingCurrency_(row, config.defaultBookingCurrency), fxRate);
  }
  return originalValue;
}

function convertAmountToGbp_(amount, currencyCode, fxRate) {
  const numericAmount = roundCurrency_(amount || 0);
  const normalizedCurrency = normalizeCurrencyCode_(currencyCode) || getReportingCurrency_();
  if (normalizedCurrency === getReportingCurrency_()) return numericAmount;
  return roundCurrency_(numericAmount * Number(fxRate || 0));
}

function getFxFallbackRateToGbp_(currencyCode) {
  const normalizedCurrency = normalizeCurrencyCode_(currencyCode) || '';
  if (!normalizedCurrency) return 0;
  return Number(FX_FALLBACK_TO_GBP[normalizedCurrency] || 0);
}

function getCachedOrFallbackFxRateToGbp_(currencyCode) {
  const normalizedCurrency = normalizeCurrencyCode_(currencyCode) || getReportingCurrency_();
  if (normalizedCurrency === getReportingCurrency_()) return 1;

  try {
    const cached = CacheService.getScriptCache().get('fx-to-gbp:' + normalizedCurrency);
    if (cached) {
      return roundFxRate_(cached);
    }
  } catch (error) {
    // Continue to static fallback. Public booking confirmation must not wait on FX services.
  }

  const fallbackRate = getFxFallbackRateToGbp_(normalizedCurrency);
  if (fallbackRate > 0) {
    return roundFxRate_(fallbackRate);
  }

  throw new Error('Cached FX rate is unavailable for ' + normalizedCurrency + ' to GBP. Please try again later or continue on WhatsApp.');
}

function getFxRateToGbp_(currencyCode) {
  const normalizedCurrency = normalizeCurrencyCode_(currencyCode) || getReportingCurrency_();
  if (normalizedCurrency === getReportingCurrency_()) return 1;

  const cache = CacheService.getScriptCache();
  const cacheKey = 'fx-to-gbp:' + normalizedCurrency;
  const cached = cache.get(cacheKey);
  if (cached) {
    return roundFxRate_(cached);
  }

  try {
    const response = UrlFetchApp.fetch(
      'https://api.frankfurter.dev/v1/latest?base=' + encodeURIComponent(normalizedCurrency) + '&symbols=GBP',
      {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          Accept: 'application/json'
        }
      }
    );
    if (response && Number(response.getResponseCode()) >= 200 && Number(response.getResponseCode()) < 300) {
      const parsed = JSON.parse(response.getContentText() || '{}');
      const liveRate = roundFxRate_(parsed && parsed.rates ? parsed.rates.GBP : 0);
      if (liveRate > 0) {
        cache.put(cacheKey, String(liveRate), FX_CACHE_TTL_SECONDS);
        return liveRate;
      }
    }
  } catch (error) {
    // Fall through to cached or static fallback.
  }

  const fallbackRate = getFxFallbackRateToGbp_(normalizedCurrency);
  if (fallbackRate > 0) {
    return roundFxRate_(fallbackRate);
  }

  throw new Error('Live FX rate is unavailable for ' + normalizedCurrency + ' to GBP. Save this booking in GBP or try again later.');
}

function getRoomTypeShortLabel_(roomTypeId, roomTypeName) {
  const normalizedRoomTypeId = String(roomTypeId || '').trim().toUpperCase();
  if (normalizedRoomTypeId === 'PEAKBAL') return 'PB';
  if (normalizedRoomTypeId === 'PEAKMV') return 'PK';
  if (normalizedRoomTypeId === 'CLASSIC') return 'CL';
  if (normalizedRoomTypeId === 'COTTAGE') return 'CT';

  const fallback = String(roomTypeName || normalizedRoomTypeId || '').trim();
  return fallback ? fallback.slice(0, 3).toUpperCase() : '';
}

function buildShortRoomLabel_(roomTypeId, roomCode, roomIdentifier) {
  const prefix = getRoomTypeShortLabel_(roomTypeId);
  const code = String(roomCode || roomIdentifier || '').trim().toUpperCase();
  return code ? (prefix ? prefix + '-' + code : code) : prefix;
}

function getBookingStatusOptions_() {
  return [
    BOOKING_STATUS_CONFIRMED,
    BOOKING_STATUS_IN_HOUSE,
    BOOKING_STATUS_CHECKED_OUT,
    BOOKING_STATUS_CANCELLED,
    BOOKING_STATUS_NO_SHOW
  ].map(function(value) {
    return { value: value, label: value };
  });
}

function generateSequentialId_(sheet, headerMap, headerKey, prefixLabel, now) {
  if (!Object.prototype.hasOwnProperty.call(headerMap, headerKey)) {
    throw new Error('Sheet is missing required header: ' + headerKey);
  }

  const prefix = prefixLabel + '-' + Utilities.formatDate(now, getScriptTimeZone_(), 'yyyyMMdd') + '-';
  const idColumn = headerMap[headerKey] + 1;
  const lastRow = sheet.getLastRow();
  let nextSequence = 1;

  if (lastRow > 1) {
    const values = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();
    values.forEach(function(row) {
      const currentValue = String(row[0] || '').trim();
      if (currentValue.indexOf(prefix) !== 0) return;

      const match = currentValue.match(/-(\d+)$/);
      if (!match) return;

      const currentSequence = Number(match[1]);
      if (currentSequence >= nextSequence) {
        nextSequence = currentSequence + 1;
      }
    });
  }

  return prefix + padNumber_(nextSequence, 3);
}

function normalizeBookingSource_(source) {
  const raw = String(source || '').trim().toLowerCase();
  if (!raw) return 'Manual';
  if (raw === 'booking.com' || raw === 'bookingcom') return 'Booking.com';
  if (raw === 'airbnb') return 'Airbnb';
  if (raw === 'website' || raw === 'direct website' || raw === 'direct-website') return 'Direct Website';
  if (raw === 'whatsapp' || raw === 'whats app') return 'WhatsApp';
  if (raw === 'phone') return 'Phone';
  if (raw === 'walk-in' || raw === 'walk in' || raw === 'walkin') return 'Walk-in';
  if (raw === 'manual') return 'Manual';
  return source;
}

function normalizeBookingStatus_(status) {
  const raw = String(status || '').trim().toLowerCase();
  if (!raw) return BOOKING_STATUS_CONFIRMED;
  if (raw === 'confirmed') return BOOKING_STATUS_CONFIRMED;
  if (raw === 'in house' || raw === 'in-house' || raw === 'inhouse') return BOOKING_STATUS_IN_HOUSE;
  if (raw === 'checked out' || raw === 'checked-out' || raw === 'checkedout') return BOOKING_STATUS_CHECKED_OUT;
  if (raw === 'cancelled' || raw === 'canceled') return BOOKING_STATUS_CANCELLED;
  if (raw === 'no show' || raw === 'no-show' || raw === 'noshow') return BOOKING_STATUS_NO_SHOW;
  return status;
}

function normalizePaymentStatus_(status, totalDue, amountPaid) {
  const raw = String(status || '').trim().toLowerCase();
  if (raw === 'pay at property' || raw === 'pay-at-property' || raw === 'pay at hotel') return 'Pay at Property';
  if (raw === 'paid') return 'Paid';
  if (raw === 'partially paid' || raw === 'partial' || raw === 'part-paid') return 'Partially Paid';
  if (raw === 'refunded') return 'Refunded';
  if (raw === 'unpaid') return 'Unpaid';

  const due = Number(totalDue || 0);
  const paid = Number(amountPaid || 0);

  if (due <= 0) {
    return paid > 0 ? 'Paid' : 'Unpaid';
  }
  if (paid <= 0) return 'Unpaid';
  if (due > 0 && paid >= due) return 'Paid';
  if (paid > 0 && due > 0) return 'Partially Paid';
  return 'Unpaid';
}

function normalizePaymentMethod_(method) {
  const raw = String(method || '').trim().toLowerCase();
  if (!raw) return '';
  if (raw === 'cash') return 'Cash';
  if (raw === 'bank transfer' || raw === 'transfer' || raw === 'bank-transfer') return 'Bank Transfer';
  if (raw === 'card' || raw === 'credit card' || raw === 'debit card') return 'Card';
  if (raw === 'booking.com virtual card' || raw === 'booking virtual card' || raw === 'virtual card') return 'Booking.com Virtual Card';
  if (raw === 'airbnb payout' || raw === 'airbnb') return 'Airbnb Payout';
  if (raw === 'online payment' || raw === 'online') return 'Online Payment';
  if (raw === 'other') return 'Other';
  return String(method || '').trim();
}

function hasRecordedPaymentMethod_(value) {
  return !!normalizePaymentMethod_(value || '');
}

function getPmsOperationalOperatorLabel_() {
  return getPmsActiveUserEmail_() || getPmsEffectiveUserEmail_() || 'PMS operator';
}

function formatOperationalAuditTimestamp_(date) {
  return Utilities.formatDate(date || new Date(), getScriptTimeZone_(), 'yyyy-MM-dd HH:mm');
}

function buildNoShowOperationalNote_(paymentDecision, operatorLabel) {
  return [
    'NO-SHOW',
    formatOperationalAuditTimestamp_(new Date()),
    operatorLabel || getPmsOperationalOperatorLabel_(),
    String(paymentDecision || '').trim()
  ].join(' | ');
}

function buildEarlyCheckoutOperationalNote_(originalCheckOut, actualCheckOut, valueAdjusted, operatorLabel) {
  return [
    'EARLY CHECK-OUT',
    'original checkout ' + formatDateKey_(originalCheckOut),
    'actual checkout ' + formatDateKey_(actualCheckOut),
    'value adjusted ' + (valueAdjusted ? 'yes' : 'no'),
    operatorLabel || getPmsOperationalOperatorLabel_()
  ].join(' | ');
}

function getInventoryEffectiveCheckOut_(row) {
  const originalCheckOut = normalizeDateInput_(row && row.check_out);
  if (!originalCheckOut) return null;

  if (normalizeBookingStatus_(row && row.status) !== BOOKING_STATUS_CHECKED_OUT) {
    return originalCheckOut;
  }

  const checkIn = normalizeDateInput_(row && row.check_in);
  const actualCheckOut = normalizeDateInput_(
    row && (row.actual_check_out_date || row.actualCheckOutDate || row.actual_checkout_date || row.actualCheckoutDate)
  );
  if (!checkIn || !actualCheckOut) return originalCheckOut;

  const actual = stripTime_(actualCheckOut);
  if (actual.getTime() <= stripTime_(checkIn).getTime()) return originalCheckOut;
  if (actual.getTime() >= stripTime_(originalCheckOut).getTime()) return originalCheckOut;
  return actual;
}

function requireCheckInReadiness_(booking) {
  if (!String(booking && (booking.room_identifier || booking.roomIdentifier) || '').trim()) {
    throw new Error('Cannot check in: assign a room first.');
  }
  if (!hasRecordedPaymentMethod_(booking && (booking.payment_method || booking.paymentMethod))) {
    throw new Error('Cannot check in: record method of payment first.');
  }
}

function isActiveFrontDeskStatus_(status) {
  return FRONT_DESK_ACTIVE_BOOKING_STATUSES.indexOf(normalizeBookingStatus_(status)) !== -1;
}

function mergeOperationalNotes_(existingText, newText) {
  const current = String(existingText || '').trim();
  const incoming = String(newText || '').trim();
  if (!current) return incoming;
  if (!incoming) return current;
  if (current.indexOf(incoming) !== -1) return current;
  return current + ' | ' + incoming;
}

function countUnassignedBookings_(rows) {
  return (rows || []).filter(function(row) {
    return !String(row.roomIdentifier || row.room_identifier || '').trim();
  }).length;
}

function countRepeatGuests_(rows) {
  return (rows || []).filter(function(row) {
    return !!row.isRepeatGuest;
  }).length;
}

function countBalanceDueBookings_(rows) {
  return (rows || []).filter(function(row) {
    return Number(row.balanceDue || row.balance_due || 0) > 0.009;
  }).length;
}

function sumBookingBalanceDue_(rows) {
  return roundCurrency_((rows || []).reduce(function(sum, row) {
    return sum + Number(row.balanceDue || row.balance_due || 0);
  }, 0));
}

function getGuestHistoryKey_(row) {
  const email = String(row.guest_email || row.guestEmail || '').trim().toLowerCase();
  if (email) return 'email:' + email;
  const phone = String(row.guest_phone || row.guestPhone || '').replace(/[^\d+]/g, '');
  if (phone) return 'phone:' + phone;
  const guestName = String(row.guest_name || row.guestName || '').trim().toLowerCase().replace(/\s+/g, ' ');
  if (guestName) return 'name:' + guestName;
  return '';
}

function buildGuestHistoryIndex_(bookingRows, options) {
  const config = options || {};
  const groups = {};
  const summaryByBookingId = {};
  const defaultBookingCurrency = config.defaultBookingCurrency || '';
  const targetLookup = {};
  const hasTargetFilter = Array.isArray(config.targetBookingIds) && config.targetBookingIds.length > 0;
  if (hasTargetFilter) {
    config.targetBookingIds.forEach(function(bookingId) {
      const key = String(bookingId || '').trim();
      if (key) targetLookup[key] = true;
    });
  }

  (bookingRows || []).forEach(function(row) {
    const bookingId = String(row.booking_id || '').trim();
    const key = getGuestHistoryKey_(row);
    if (!bookingId || !key) return;
    if (!groups[key]) groups[key] = [];
    groups[key].push(row);
  });

  Object.keys(groups).forEach(function(key) {
    const group = groups[key].map(function(row) {
      const checkIn = normalizeDateInput_(row.check_in);
      const checkOut = normalizeDateInput_(row.check_out);
      return {
        row: row,
        bookingId: String(row.booking_id || '').trim(),
        checkIn: checkIn,
        checkInMs: checkIn ? stripTime_(checkIn).getTime() : 0,
        checkOut: checkOut,
        checkOutMs: checkOut ? stripTime_(checkOut).getTime() : 0,
        counted: isRevenueCountedStatus_(normalizeBookingStatus_(row.status)),
        guestPreference: String(row.guest_preferences || '').trim()
      };
    }).sort(function(a, b) {
      if (a.checkInMs !== b.checkInMs) return a.checkInMs - b.checkInMs;
      if (a.checkOutMs !== b.checkOutMs) return b.checkOutMs - a.checkOutMs;
      return a.bookingId.localeCompare(b.bookingId);
    });

    let revenueSoFar = [];
    let index = 0;
    while (index < group.length) {
      const checkInMs = group[index].checkInMs;
      const batch = [];
      while (index < group.length && group[index].checkInMs === checkInMs) {
        batch.push(group[index]);
        index += 1;
      }

      batch.forEach(function(item) {
        if (hasTargetFilter && !targetLookup[item.bookingId]) return;
        const previousRows = item.checkIn ? revenueSoFar : [];
        const lastPreference = item.guestPreference || (previousRows.filter(function(previous) {
          return previous.guestPreference;
        })[0] || {}).guestPreference || '';

        summaryByBookingId[item.bookingId] = {
        isRepeatGuest: previousRows.length > 0,
        repeatStayCount: previousRows.length,
        lastStayCheckOut: previousRows.length && previousRows[0].checkOut
          ? formatDateKey_(previousRows[0].checkOut)
          : '',
        lastPreference: lastPreference,
        previousStays: previousRows.slice(0, 5).map(function(previous) {
          return {
            bookingId: previous.bookingId,
            checkIn: previous.checkIn ? formatDateKey_(previous.checkIn) : '',
            checkOut: previous.checkOut ? formatDateKey_(previous.checkOut) : '',
            roomTypeName: String(previous.row.room_type_name || previous.row.room_type_id || '').trim(),
            roomIdentifier: String(previous.row.room_identifier || '').trim(),
            status: normalizeBookingStatus_(previous.row.status),
            source: normalizeBookingSource_(previous.row.source || ''),
            bookingValue: getOperationalBookingValueGbp_(previous.row, {
              defaultBookingCurrency: defaultBookingCurrency
            }),
            guestPreferences: previous.guestPreference,
            notes: String(previous.row.notes || '').trim()
          };
        })
      };
    });

      const countedBatch = batch.filter(function(item) {
        return item.counted && item.checkIn;
      });
      if (countedBatch.length) {
        revenueSoFar = revenueSoFar.concat(countedBatch).sort(function(a, b) {
          if (a.checkOutMs !== b.checkOutMs) return b.checkOutMs - a.checkOutMs;
          return b.checkInMs - a.checkInMs;
        });
      }
    }
  });

  return summaryByBookingId;
}

function buildBookingPrepSummary_(roomIdentifier, balanceDue, guestPreferences, paymentNotes, notes) {
  const items = [];
  if (!String(roomIdentifier || '').trim()) items.push('Room not assigned');
  if (Number(balanceDue || 0) > 0.009) items.push('Balance due');
  if (String(guestPreferences || '').trim()) items.push('Preference recorded');
  if (String(notes || '').trim()) items.push('Special note');
  if (String(paymentNotes || '').trim()) items.push('Payment note');
  return items.join(' | ');
}

function buildPortfolioInsights_(bookings, startDate, endDate, options) {
  const config = options || {};
  const valueContext = {
    defaultBookingCurrency: config.defaultBookingCurrency || ''
  };
  const start = stripTime_(startDate);
  const end = stripTime_(endDate);
  const revenueRows = (bookings || []).filter(function(row) {
    const checkIn = normalizeDateInput_(row.check_in);
    return checkIn &&
      checkIn.getTime() >= start.getTime() &&
      checkIn.getTime() <= end.getTime() &&
      isRevenueCountedStatus_(row.status);
  });
  const totalRevenue = revenueRows.reduce(function(sum, row) {
    return sum + getOperationalBookingValueGbp_(row, valueContext);
  }, 0);
  const directRows = revenueRows.filter(function(row) {
    return normalizeBookingSource_(row.source || '') === 'Direct Website';
  });
  const directRevenue = directRows.reduce(function(sum, row) {
    return sum + getOperationalBookingValueGbp_(row, valueContext);
  }, 0);
  const cancellationCount = (bookings || []).filter(function(row) {
    if (normalizeBookingStatus_(row.status) !== BOOKING_STATUS_CANCELLED) return false;
    const cancelledAt = normalizeDateInput_(row.cancelled_at || row.created_at);
    return cancelledAt && cancelledAt.getTime() >= start.getTime() && cancelledAt.getTime() <= end.getTime();
  }).length;

  const leadTimes = revenueRows.map(function(row) {
    const createdAt = normalizeDateInput_(row.created_at);
    const checkIn = normalizeDateInput_(row.check_in);
    if (!createdAt || !checkIn) return null;
    return Math.max(0, Math.round((stripTime_(checkIn).getTime() - stripTime_(createdAt).getTime()) / 86400000));
  }).filter(function(value) {
    return value != null;
  });

  return {
    directWebsiteBookings: directRows.length,
    directWebsiteSharePct: safeDivide_(directRevenue, totalRevenue),
    cancellationsCount: cancellationCount,
    averageLeadTimeDays: leadTimes.length ? roundCurrency_(leadTimes.reduce(function(sum, value) {
      return sum + value;
    }, 0) / leadTimes.length) : 0
  };
}

function buildOccupancyTrend_(dailyStatsRows, startDate, days) {
  const rows = [];
  const safeDays = Math.max(1, Number(days || 7));

  for (let offset = 0; offset < safeDays; offset++) {
    const currentDate = addDays_(startDate, offset);
    const metrics = getOccupancyMetricsForDate_(dailyStatsRows, currentDate);
    rows.push({
      date: formatDateKey_(currentDate),
      occupancyPct: metrics.occupancyPct,
      soldRooms: metrics.roomNights,
      availableRooms: metrics.availableRoomNights
    });
  }

  return rows;
}

function getOccupancyMetricsForDate_(dailyStatsRows, targetDate) {
  const targetKey = formatDateKey_(targetDate);
  let roomNights = 0;
  let availableRoomNights = 0;

  (dailyStatsRows || []).forEach(function(row) {
    if (String(row.stay_date || '').trim() !== targetKey) return;
    roomNights += Number(row.room_nights || 0);
    availableRoomNights += Number(row.rooms_available_to_sell || 0);
  });

  return {
    roomNights: roomNights,
    availableRoomNights: availableRoomNights,
    occupancyPct: safeDivide_(roomNights, availableRoomNights)
  };
}

function mapRequestSourceToBookingSource_(requestSource) {
  const raw = String(requestSource || '').trim().toLowerCase();
  if (raw === 'website') return 'Direct Website';
  if (raw === 'whatsapp') return 'WhatsApp';
  if (raw === 'phone') return 'Phone';
  return 'Manual';
}

function toBoolean_(value) {
  if (typeof value === 'boolean') return value;
  const raw = String(value || '').trim().toLowerCase();
  return raw === 'true' || raw === 'yes' || raw === '1';
}

function validateMoneyAmountInput_(value, label, defaultValue) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return roundCurrency_(defaultValue || 0);
  const number = Number(raw);
  if (!isFinite(number)) {
    throw new Error((label || 'Money amount') + ' must be a valid number.');
  }
  if (number < 0) {
    throw new Error((label || 'Money amount') + ' cannot be negative.');
  }
  return roundCurrency_(number);
}

function assertValidMoneyAmount_(value, label) {
  const number = Number(value);
  if (!isFinite(number)) {
    throw new Error((label || 'Money amount') + ' must be a valid number.');
  }
  if (number < 0) {
    throw new Error((label || 'Money amount') + ' cannot be negative.');
  }
  return true;
}

function roundCurrency_(value) {
  const number = Number(value || 0);
  if (!isFinite(number)) return 0;
  return Math.round(number * 100) / 100;
}

function calculateBalanceDue_(bookingValue, amountPaid) {
  return roundCurrency_(Number(bookingValue || 0) - Number(amountPaid || 0));
}

function safeDivide_(numerator, denominator) {
  const top = Number(numerator || 0);
  const bottom = Number(denominator || 0);
  return bottom ? top / bottom : 0;
}

function addDays_(date, days) {
  const result = new Date(stripTime_(date));
  result.setDate(result.getDate() + Number(days || 0));
  return result;
}

function enumerateDatesInclusive_(startDate, endDate) {
  const dates = [];
  for (let cursor = stripTime_(startDate); cursor.getTime() <= stripTime_(endDate).getTime(); cursor = addDays_(cursor, 1)) {
    dates.push(new Date(cursor));
  }
  return dates;
}

function shiftDateYears_(date, years) {
  return new Date(date.getFullYear() + Number(years || 0), date.getMonth(), date.getDate());
}

function startOfMonth_(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function startOfYear_(date) {
  return new Date(date.getFullYear(), 0, 1);
}

function formatMonthKey_(date) {
  return Utilities.formatDate(stripTime_(date), getScriptTimeZone_(), 'yyyy-MM');
}

function formatMoneyValue_(amount, currency) {
  return String(currency || getReportingCurrency_()) + ' ' + Utilities.formatString('%.2f', Number(roundCurrency_(amount || 0)));
}

function getSettingValue_(settings, keys, fallbackValue) {
  const options = Array.isArray(keys) ? keys : [keys];
  for (let i = 0; i < options.length; i++) {
    const key = String(options[i] || '').trim();
    if (!key) continue;
    const value = settings && Object.prototype.hasOwnProperty.call(settings, key) ? settings[key] : '';
    if (String(value || '').trim()) {
      return String(value).trim();
    }
  }
  return fallbackValue;
}

function getWebsiteContactDetails_(prefetchedSettings) {
  let settings = prefetchedSettings || {};

  if (!prefetchedSettings) {
    try {
      settings = getSettings() || {};
    } catch (error) {
      settings = {};
    }
  }

  const propertyName = getSettingValue_(settings, ['property_name', 'propertyName', 'hotel_name', 'business_name'], DEFAULT_PROPERTY_NAME);
  const propertyAddress = getSettingValue_(settings, ['property_address', 'propertyAddress', 'address'], DEFAULT_PROPERTY_ADDRESS);
  const phone = getSettingValue_(settings, ['contact_phone', 'phone', 'property_phone'], DEFAULT_PROPERTY_PHONE);
  const whatsapp = getSettingValue_(settings, ['whatsapp_number', 'whatsapp', 'contact_whatsapp'], phone || DEFAULT_PROPERTY_PHONE);
  const email = getSettingValue_(settings, ['contact_email', 'email', 'property_email'], DEFAULT_PROPERTY_EMAIL);
  const notificationEmail = getSettingValue_(settings, ['booking_notification_email', 'internal_notification_email', 'notification_email', 'reservations_email'], email);

  return {
    propertyName: propertyName,
    propertyAddress: propertyAddress,
    phone: phone,
    whatsapp: whatsapp,
    email: email,
    notificationEmail: notificationEmail
  };
}

function isRevenueCountedStatus_(status) {
  return BOOKING_STATUSES_COUNTED.indexOf(normalizeBookingStatus_(status)) !== -1;
}

function isPendingRequestStatus_(status) {
  return REQUEST_STATUS_PENDING_VALUES.indexOf(String(status || '').trim()) !== -1;
}

function isActiveRoomType_(roomTypeRow) {
  const activeFlag = String(roomTypeRow.active || '').trim().toLowerCase();
  return !activeFlag || activeFlag === 'yes' || activeFlag === 'true' || activeFlag === 'active';
}

function getRemainingInventoryForStay_(checkIn, checkOut, roomTypeId) {
  const options = arguments[3] || {};
  const inventoryTotal = getRoomInventory_(roomTypeId);
  const excludeBookingId = String(options.excludeBookingId || '').trim();
  const bookingRows = excludeBookingId
    ? getSheetObjects_(SHEET_NAMES.BOOKINGS).filter(function(row) {
        return String(row.booking_id || '').trim() !== excludeBookingId;
      })
    : null;
  const bookingCountMap = bookingRows ? buildConfirmedBookingRoomCountMap_(bookingRows).byDateRoom : null;
  const blockedCountMap = bookingRows ? buildBlockedDateQtyMap_(getSheetObjects_(SHEET_NAMES.BLOCKED_DATES)).byDateRoom : null;
  const commercialControls = getCommercialControlRows_({ activeOnly: true, roomTypeId: roomTypeId });
  let remainingMin = Number.POSITIVE_INFINITY;

  enumerateStayDates_(checkIn, checkOut).forEach(function(date) {
    const sellableInventory = getSellableInventoryTotalForDate_(date, roomTypeId, commercialControls, inventoryTotal);
    const remaining = sellableInventory -
      getConfirmedBookingsCount_(date, roomTypeId, bookingCountMap) -
      getBlockedDatesCount_(date, roomTypeId, blockedCountMap);
    if (remaining < remainingMin) {
      remainingMin = remaining;
    }
  });

  return isFinite(remainingMin) ? remainingMin : inventoryTotal;
}

function syncRequestAsConverted_(requestId, details) {
  const spreadsheet = getSpreadsheet_();
  const requestSheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.REQUESTS, REQUESTS_HEADERS);
  const rowNumber = findRowNumberByHeaderValue_(requestSheet, 'request_id', requestId);

  if (!rowNumber) {
    return false;
  }

  updateObjectRowBulk_(requestSheet, rowNumber, {
    request_status: REQUEST_STATUS_CONVERTED,
    booking_id: details.bookingId || '',
    converted_to_booking_at: details.convertedAt || new Date(),
    conversion_status: REQUEST_STATUS_CONVERTED,
    assigned_source: details.source || '',
    assigned_booking_value: details.bookingValue || ''
  });

  return true;
}

function rebuildBookingNights_(spreadsheet) {
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.BOOKING_NIGHTS, BOOKING_NIGHTS_HEADERS);
  const bookings = getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const defaultCurrency = getReportingCurrency_();
  const rows = [BOOKING_NIGHTS_HEADERS];

  bookings.forEach(function(booking) {
    const status = normalizeBookingStatus_(booking.status);
    if (!isRevenueCountedStatus_(status)) return;

    const checkIn = normalizeDateInput_(booking.check_in);
    const checkOut = getInventoryEffectiveCheckOut_(booking);
    if (!checkIn || !checkOut || checkOut.getTime() <= checkIn.getTime()) return;

    const stayDates = enumerateStayDates_(checkIn, checkOut);
    if (!stayDates.length) return;

    const qtyRooms = Math.max(1, Number(booking.qty_rooms || 1));
    const bookingValue = getOperationalBookingValueGbp_(booking);
    const effectiveNights = stayDates.length;
    const nightlyRoomRevenue = roundCurrency_(safeDivide_(bookingValue, effectiveNights));

    stayDates.forEach(function(stayDate) {
      rows.push([
        formatDateKey_(stayDate),
        formatMonthKey_(stayDate),
        booking.booking_id || '',
        booking.request_id || '',
        booking.created_at || '',
        booking.source || '',
        booking.source_detail || '',
        status,
        booking.guest_name || '',
        booking.room_type_id || '',
        booking.room_type_name || '',
        qtyRooms,
        qtyRooms,
        nightlyRoomRevenue,
        bookingValue,
        booking.check_in || '',
        booking.check_out || '',
        effectiveNights,
        defaultCurrency
      ]);
    });
  });

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  styleHeaderRow_(sheet, 1, rows[0].length);

  return { ok: true, rowsWritten: rows.length - 1 };
}

function rebuildDailyStats_(spreadsheet) {
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.DAILY_STATS, DAILY_STATS_HEADERS);
  const bookingNightRows = getSheetObjects_(SHEET_NAMES.BOOKING_NIGHTS);
  const blockedRows = getSheetObjects_(SHEET_NAMES.BLOCKED_DATES);
  const roomTypes = getActiveRoomTypeCatalog_();
  const today = stripTime_(new Date());
  let startDate = new Date(today.getFullYear() - 1, 0, 1);
  let endDate = addDays_(today, OTB_SNAPSHOT_DAYS_AHEAD);
  const bookingNightStats = buildBookingNightStatsMap_(bookingNightRows);
  const blockedDateMap = buildBlockedDateQtyMap_(blockedRows);
  const rows = [DAILY_STATS_HEADERS];

  if (bookingNightStats.minDate && bookingNightStats.minDate.getTime() < startDate.getTime()) {
    startDate = bookingNightStats.minDate;
  }
  if (bookingNightStats.maxDate && bookingNightStats.maxDate.getTime() > endDate.getTime()) {
    endDate = bookingNightStats.maxDate;
  }
  if (blockedDateMap.minDate && blockedDateMap.minDate.getTime() < startDate.getTime()) {
    startDate = blockedDateMap.minDate;
  }
  if (blockedDateMap.maxDate && blockedDateMap.maxDate.getTime() > endDate.getTime()) {
    endDate = blockedDateMap.maxDate;
  }

  for (let cursor = new Date(startDate); cursor.getTime() <= endDate.getTime(); cursor = addDays_(cursor, 1)) {
    const dateKey = formatDateKey_(cursor);
    roomTypes.forEach(function(roomType) {
      const roomTypeId = String(roomType.roomTypeId || '').trim();
      const roomTypeName = String(roomType.roomTypeName || roomTypeId).trim();
      const inventoryTotal = Number(roomType.inventoryTotal || 0);
      const key = dateKey + '|' + roomTypeId;
      const blockedRooms = Number(blockedDateMap.byDateRoom[key] || 0);
      const roomsAvailableToSell = Math.max(0, inventoryTotal - blockedRooms);
      const stats = bookingNightStats.byDateRoom[key] || { roomsSold: 0, roomNights: 0, roomRevenue: 0, bookingsCount: 0 };
      const roomRevenue = roundCurrency_(stats.roomRevenue);
      const occupancyPct = safeDivide_(stats.roomsSold, roomsAvailableToSell);
      const adr = safeDivide_(roomRevenue, stats.roomNights);
      const revpar = safeDivide_(roomRevenue, roomsAvailableToSell);

      rows.push([
        dateKey,
        formatMonthKey_(cursor),
        cursor.getFullYear(),
        cursor.getMonth() + 1,
        roomTypeId,
        roomTypeName,
        inventoryTotal,
        blockedRooms,
        roomsAvailableToSell,
        stats.roomsSold,
        stats.roomNights,
        stats.bookingsCount,
        roomRevenue,
        occupancyPct,
        adr,
        revpar
      ]);
    });
  }

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  styleHeaderRow_(sheet, 1, rows[0].length);

  return { ok: true, rowsWritten: rows.length - 1 };
}

function rebuildOtbSnapshot_(spreadsheet) {
  const sheet = ensureSheetWithHeaders_(spreadsheet, SHEET_NAMES.OTB_SNAPSHOTS, OTB_SNAPSHOTS_HEADERS);
  const dailyStatsRows = getSheetObjects_(SHEET_NAMES.DAILY_STATS);
  const snapshotDate = stripTime_(new Date());
  const snapshotKey = formatDateKey_(snapshotDate);
  const endDate = addDays_(snapshotDate, OTB_SNAPSHOT_DAYS_AHEAD);
  const existingValues = sheet.getDataRange().getValues();
  const rows = [OTB_SNAPSHOTS_HEADERS];

  if (existingValues.length > 1) {
    existingValues.slice(1).forEach(function(row) {
      if (String(row[0] || '').trim() !== snapshotKey) {
        rows.push(row);
      }
    });
  }

  dailyStatsRows.forEach(function(row) {
    const stayDate = normalizeDateInput_(row.stay_date);
    if (!stayDate) return;
    if (stayDate.getTime() < snapshotDate.getTime() || stayDate.getTime() > endDate.getTime()) return;

    rows.push([
      snapshotKey,
      row.stay_date,
      row.stay_month,
      row.room_type_id,
      row.room_type_name,
      Number(row.rooms_available_to_sell || 0),
      Number(row.rooms_sold || 0),
      roundCurrency_(row.room_revenue || 0),
      Number(row.occupancy_pct || 0),
      roundCurrency_(row.adr || 0),
      roundCurrency_(row.revpar || 0)
    ]);
  });

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  styleHeaderRow_(sheet, 1, rows[0].length);

  return { ok: true, rowsWritten: rows.length - 1, snapshotDate: snapshotKey };
}

function rebuildDashboard_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SHEET_NAMES.DASHBOARD) || spreadsheet.insertSheet(SHEET_NAMES.DASHBOARD);
  const previousSearch = String(sheet.getRange(DASHBOARD_SEARCH_CELL).getValue() || '').trim();
  const today = stripTime_(new Date());
  const monthStart = startOfMonth_(today);
  const yearStart = startOfYear_(today);
  const lyToday = shiftDateYears_(today, -1);
  const lyMonthStart = startOfMonth_(lyToday);
  const lyYearStart = startOfYear_(lyToday);
  const bookings = getSheetObjects_(SHEET_NAMES.BOOKINGS);
  const requests = getSheetObjects_(SHEET_NAMES.REQUESTS);
  const bookingNights = getSheetObjects_(SHEET_NAMES.BOOKING_NIGHTS);
  const dailyStats = getSheetObjects_(SHEET_NAMES.DAILY_STATS);
  const snapshots = getSheetObjects_(SHEET_NAMES.OTB_SNAPSHOTS);
  const currency = getReportingCurrency_();
  const arrivals = getArrivalsOrDepartures_(bookings, today, 'check_in');
  const departures = getArrivalsOrDepartures_(bookings, today, 'check_out');
  const inHouse = getInHouseBookings_(bookings, today);
  const pendingRequests = requests.filter(function(row) {
    return isPendingRequestStatus_(row.request_status);
  });
  const confirmedBookings = bookings.filter(function(row) {
    return normalizeBookingStatus_(row.status) === BOOKING_STATUS_CONFIRMED;
  });
  const cancelledBookingsMtd = bookings.filter(function(row) {
    const status = normalizeBookingStatus_(row.status);
    if (status !== BOOKING_STATUS_CANCELLED) return false;
    const cancelledAt = normalizeDateInput_(row.cancelled_at || row.created_at);
    return cancelledAt && cancelledAt.getTime() >= monthStart.getTime() && cancelledAt.getTime() <= today.getTime();
  });
  const mtdMetrics = getPeriodMetrics_(dailyStats, bookingNights, bookings, monthStart, today);
  const ytdMetrics = getPeriodMetrics_(dailyStats, bookingNights, bookings, yearStart, today);
  const lyMtdMetrics = getPeriodMetrics_(dailyStats, bookingNights, bookings, lyMonthStart, lyToday);
  const lyYtdMetrics = getPeriodMetrics_(dailyStats, bookingNights, bookings, lyYearStart, lyToday);
  const otb30 = getPeriodMetrics_(dailyStats, bookingNights, bookings, today, addDays_(today, 29));
  const otb90 = getPeriodMetrics_(dailyStats, bookingNights, bookings, today, addDays_(today, 89));
  const paceLy = getPaceComparison_(snapshots, today, 90);
  const pickup = getPickupComparison_(snapshots, today, 7, 90);
  const sourceSummary = summariseNightRowsByField_(bookingNights, monthStart, today, 'source');
  const roomTypeSummary = summariseNightRowsByField_(bookingNights, monthStart, today, 'room_type_name');
  const searchResults = buildDashboardSearchResults_(bookings, requests, previousSearch);

  sheet.clearContents().clearFormats();
  sheet.setFrozenRows(3);
  sheet.getRange('A1').setValue('Roza PMS Dashboard').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A2').setValue('Last refreshed');
  sheet.getRange('B2').setValue(new Date());
  sheet.getRange('A3').setValue('Search by request ID / booking ID / guest name');
  sheet.getRange(DASHBOARD_SEARCH_CELL).setValue(previousSearch);
  sheet.getRange('C3').setValue('Edit B3 and run Refresh Dashboard Only');

  let row = 5;

  row = writeKeyValueBlock_(sheet, row, 'Today', [
    ['Arrivals today', arrivals.length],
    ['Departures today', departures.length],
    ['In-house guests', inHouse.length],
    ['Pending requests', pendingRequests.length],
    ['Confirmed bookings', confirmedBookings.length],
    ['Cancelled bookings MTD', cancelledBookingsMtd.length]
  ]);

  row = writeTableBlock_(sheet, row, 'Performance Summary', ['Period', 'Room Revenue', 'Room Nights', 'Occupancy %', 'ADR', 'RevPAR', 'ALOS'], [
    ['MTD', formatMoney_(mtdMetrics.roomRevenue, currency), mtdMetrics.roomNights, formatPercent_(mtdMetrics.occupancyPct), formatMoney_(mtdMetrics.adr, currency), formatMoney_(mtdMetrics.revpar, currency), roundMetric_(mtdMetrics.alos)],
    ['MTD LY', formatMoney_(lyMtdMetrics.roomRevenue, currency), lyMtdMetrics.roomNights, formatPercent_(lyMtdMetrics.occupancyPct), formatMoney_(lyMtdMetrics.adr, currency), formatMoney_(lyMtdMetrics.revpar, currency), roundMetric_(lyMtdMetrics.alos)],
    ['YTD', formatMoney_(ytdMetrics.roomRevenue, currency), ytdMetrics.roomNights, formatPercent_(ytdMetrics.occupancyPct), formatMoney_(ytdMetrics.adr, currency), formatMoney_(ytdMetrics.revpar, currency), roundMetric_(ytdMetrics.alos)],
    ['YTD LY', formatMoney_(lyYtdMetrics.roomRevenue, currency), lyYtdMetrics.roomNights, formatPercent_(lyYtdMetrics.occupancyPct), formatMoney_(lyYtdMetrics.adr, currency), formatMoney_(lyYtdMetrics.revpar, currency), roundMetric_(lyYtdMetrics.alos)],
    ['OTB next 30', formatMoney_(otb30.roomRevenue, currency), otb30.roomNights, formatPercent_(otb30.occupancyPct), formatMoney_(otb30.adr, currency), formatMoney_(otb30.revpar, currency), roundMetric_(otb30.alos)],
    ['OTB next 90', formatMoney_(otb90.roomRevenue, currency), otb90.roomNights, formatPercent_(otb90.occupancyPct), formatMoney_(otb90.adr, currency), formatMoney_(otb90.revpar, currency), roundMetric_(otb90.alos)]
  ]);

  row = writeTableBlock_(sheet, row, 'Pace And Pickup', ['Metric', 'Value'], [
    ['Pace vs same time last year (next 90 days revenue)', paceLy.available ? formatMoney_(paceLy.currentRevenue - paceLy.lastYearRevenue, currency) : 'Not enough snapshot history yet'],
    ['Pace vs same time last year (next 90 days room nights)', paceLy.available ? (paceLy.currentRoomNights - paceLy.lastYearRoomNights) : 'Not enough snapshot history yet'],
    ['Pickup vs 7 days ago (next 90 days revenue)', pickup.available ? formatMoney_(pickup.currentRevenue - pickup.previousRevenue, currency) : 'Not enough snapshot history yet'],
    ['Pickup vs 7 days ago (next 90 days room nights)', pickup.available ? (pickup.currentRoomNights - pickup.previousRoomNights) : 'Not enough snapshot history yet']
  ]);

  row = writeTableBlock_(sheet, row, 'Revenue By Source MTD', ['Source', 'Room Revenue', 'Room Nights'], sourceSummary);
  row = writeTableBlock_(sheet, row, 'Revenue By Room Type MTD', ['Room Type', 'Room Revenue', 'Room Nights'], roomTypeSummary);

  row = writeTableBlock_(sheet, row, 'Arrivals Today', ['Booking ID', 'Guest', 'Room Type', 'Source', 'Rooms', 'Guests', 'Value'], mapBookingsForDashboard_(arrivals, currency));
  row = writeTableBlock_(sheet, row, 'Departures Today', ['Booking ID', 'Guest', 'Room Type', 'Source', 'Rooms', 'Guests', 'Value'], mapBookingsForDashboard_(departures, currency));
  row = writeTableBlock_(sheet, row, 'In-House Guests', ['Booking ID', 'Guest', 'Room Type', 'Source', 'Rooms', 'Guests', 'Value'], mapBookingsForDashboard_(inHouse, currency));
  row = writeTableBlock_(sheet, row, 'Pending Requests', ['Request ID', 'Guest', 'Check-in', 'Check-out', 'Room Type', 'Source', 'Status'], mapRequestsForDashboard_(pendingRequests));
  row = writeTableBlock_(sheet, row, 'Search Bookings', ['Booking ID', 'Request ID', 'Guest', 'Check-in', 'Check-out', 'Status', 'Source'], searchResults.bookings);
  row = writeTableBlock_(sheet, row, 'Search Requests', ['Request ID', 'Guest', 'Check-in', 'Check-out', 'Status', 'Source', 'Booking ID'], searchResults.requests);

  sheet.autoResizeColumns(1, Math.max(7, sheet.getLastColumn()));

  return {
    ok: true,
    searchTerm: previousSearch,
    arrivalsToday: arrivals.length,
    departuresToday: departures.length,
    inHouseGuests: inHouse.length,
    pendingRequests: pendingRequests.length
  };
}

function getPeriodMetrics_(dailyStatsRows, bookingNightRows, bookings, startDate, endDate) {
  const start = stripTime_(startDate);
  const end = stripTime_(endDate);
  let roomRevenue = 0;
  let roomNights = 0;
  let availableRoomNights = 0;
  let bookingCount = 0;
  let totalBookingNights = 0;

  dailyStatsRows.forEach(function(row) {
    const stayDate = normalizeDateInput_(row.stay_date);
    if (!stayDate || stayDate.getTime() < start.getTime() || stayDate.getTime() > end.getTime()) return;
    roomRevenue += Number(row.room_revenue || 0);
    roomNights += Number(row.room_nights || 0);
    availableRoomNights += Number(row.rooms_available_to_sell || 0);
  });

  bookings.forEach(function(row) {
    const checkIn = normalizeDateInput_(row.check_in);
    const checkOut = getInventoryEffectiveCheckOut_(row);
    const status = normalizeBookingStatus_(row.status);
    if (!checkIn || !checkOut || !isRevenueCountedStatus_(status)) return;
    if (checkIn.getTime() < start.getTime() || checkIn.getTime() > end.getTime()) return;
    const effectiveStayDates = enumerateStayDates_(checkIn, checkOut);
    if (!effectiveStayDates.length) return;
    bookingCount += 1;
    totalBookingNights += effectiveStayDates.length;
  });

  return {
    roomRevenue: roundCurrency_(roomRevenue),
    roomNights: roomNights,
    availableRoomNights: availableRoomNights,
    occupancyPct: safeDivide_(roomNights, availableRoomNights),
    adr: safeDivide_(roomRevenue, roomNights),
    revpar: safeDivide_(roomRevenue, availableRoomNights),
    bookingsCount: bookingCount,
    alos: safeDivide_(totalBookingNights, bookingCount)
  };
}

function summariseNightRowsByField_(bookingNightRows, startDate, endDate, fieldName) {
  const start = stripTime_(startDate);
  const end = stripTime_(endDate);
  const buckets = {};

  bookingNightRows.forEach(function(row) {
    const stayDate = normalizeDateInput_(row.stay_date);
    if (!stayDate || stayDate.getTime() < start.getTime() || stayDate.getTime() > end.getTime()) return;

    const key = String(row[fieldName] || 'Unassigned').trim() || 'Unassigned';
    if (!buckets[key]) {
      buckets[key] = { revenue: 0, roomNights: 0 };
    }
    buckets[key].revenue += Number(row.nightly_room_revenue || 0);
    buckets[key].roomNights += Number(row.room_nights || 0);
  });

  const rows = Object.keys(buckets).sort().map(function(key) {
    return [key, roundCurrency_(buckets[key].revenue), buckets[key].roomNights];
  });

  return rows.length ? rows : [['No data yet', '', '']];
}

function getArrivalsOrDepartures_(bookings, targetDate, fieldName) {
  const target = stripTime_(targetDate).getTime();
  return bookings.filter(function(row) {
    const status = normalizeBookingStatus_(row.status);
    const dateValue = normalizeDateInput_(row[fieldName]);
    if (!dateValue) return false;
    if (status === BOOKING_STATUS_CANCELLED || status === BOOKING_STATUS_NO_SHOW) return false;
    return stripTime_(dateValue).getTime() === target;
  });
}

function getInHouseBookings_(bookings, targetDate) {
  const target = stripTime_(targetDate).getTime();
  return bookings.filter(function(row) {
    const status = normalizeBookingStatus_(row.status);
    if (status === BOOKING_STATUS_CANCELLED || status === BOOKING_STATUS_NO_SHOW) return false;
    const checkIn = normalizeDateInput_(row.check_in);
    const checkOut = normalizeDateInput_(row.check_out);
    if (!checkIn || !checkOut) return false;
    return target >= stripTime_(checkIn).getTime() && target < stripTime_(checkOut).getTime();
  });
}

function getPaceComparison_(snapshotRows, today, daysAhead) {
  const currentSnapshot = formatDateKey_(today);
  const lastYearSnapshot = formatDateKey_(shiftDateYears_(today, -1));
  const currentEnd = addDays_(today, daysAhead - 1);
  const lastYearStart = shiftDateYears_(today, -1);
  const lastYearEnd = addDays_(lastYearStart, daysAhead - 1);
  const current = sumSnapshotMetrics_(snapshotRows, currentSnapshot, today, currentEnd);
  const lastYear = sumSnapshotMetrics_(snapshotRows, lastYearSnapshot, lastYearStart, lastYearEnd);

  return {
    available: current.found && lastYear.found,
    currentRevenue: current.roomRevenue,
    currentRoomNights: current.roomNights,
    lastYearRevenue: lastYear.roomRevenue,
    lastYearRoomNights: lastYear.roomNights
  };
}

function getPickupComparison_(snapshotRows, today, daysBack, daysAhead) {
  const currentSnapshot = formatDateKey_(today);
  const previousSnapshotDate = addDays_(today, -Number(daysBack || 0));
  const previousSnapshot = formatDateKey_(previousSnapshotDate);
  const currentEnd = addDays_(today, daysAhead - 1);
  const previousEnd = addDays_(previousSnapshotDate, daysAhead - 1);
  const current = sumSnapshotMetrics_(snapshotRows, currentSnapshot, today, currentEnd);
  const previous = sumSnapshotMetrics_(snapshotRows, previousSnapshot, previousSnapshotDate, previousEnd);

  return {
    available: current.found && previous.found,
    currentRevenue: current.roomRevenue,
    currentRoomNights: current.roomNights,
    previousRevenue: previous.roomRevenue,
    previousRoomNights: previous.roomNights
  };
}

function sumSnapshotMetrics_(snapshotRows, snapshotKey, stayStart, stayEnd) {
  let roomRevenue = 0;
  let roomNights = 0;
  let found = false;
  const start = stripTime_(stayStart);
  const end = stripTime_(stayEnd);

  snapshotRows.forEach(function(row) {
    if (String(row.snapshot_date || '').trim() !== snapshotKey) return;
    const stayDate = normalizeDateInput_(row.stay_date);
    if (!stayDate || stayDate.getTime() < start.getTime() || stayDate.getTime() > end.getTime()) return;
    found = true;
    roomRevenue += Number(row.room_revenue_otb || 0);
    roomNights += Number(row.rooms_sold_otb || 0);
  });

  return {
    found: found,
    roomRevenue: roundCurrency_(roomRevenue),
    roomNights: roomNights
  };
}

function buildDashboardSearchResults_(bookings, requests, term) {
  const raw = String(term || '').trim().toLowerCase();
  if (!raw) {
    return {
      bookings: [['Enter a search term in B3 and refresh', '', '', '', '', '', '']],
      requests: [['Enter a search term in B3 and refresh', '', '', '', '', '', '']]
    };
  }

  const bookingMatches = bookings.filter(function(row) {
    return matchesSearchTerm_(row.booking_id, raw) || matchesSearchTerm_(row.request_id, raw) || matchesSearchTerm_(row.guest_name, raw);
  }).slice(0, 10).map(function(row) {
    return [
      row.booking_id || '',
      row.request_id || '',
      row.guest_name || '',
      row.check_in || '',
      row.check_out || '',
      row.status || '',
      row.source || ''
    ];
  });

  const requestMatches = requests.filter(function(row) {
    return matchesSearchTerm_(row.request_id, raw) || matchesSearchTerm_(row.booking_id, raw) || matchesSearchTerm_(row.guest_name, raw);
  }).slice(0, 10).map(function(row) {
    return [
      row.request_id || '',
      row.guest_name || '',
      row.check_in || '',
      row.check_out || '',
      row.request_status || '',
      row.request_source || '',
      row.booking_id || ''
    ];
  });

  return {
    bookings: bookingMatches.length ? bookingMatches : [['No matching bookings', '', '', '', '', '', '']],
    requests: requestMatches.length ? requestMatches : [['No matching requests', '', '', '', '', '', '']]
  };
}

function mapBookingsForDashboard_(rows, currency) {
  return rows.length ? rows.map(function(row) {
    return [
      row.booking_id || '',
      row.guest_name || '',
      row.room_type_name || row.room_type_id || '',
      row.source || '',
      Number(row.qty_rooms || 1),
      Number(row.guests || 1),
      formatMoney_(getOperationalBookingValueGbp_(row), currency)
    ];
  }) : [['No rows', '', '', '', '', '', '']];
}

function mapRequestsForDashboard_(rows) {
  return rows.length ? rows.map(function(row) {
    return [
      row.request_id || '',
      row.guest_name || '',
      row.check_in || '',
      row.check_out || '',
      row.room_type_name || row.room_type_id || '',
      row.request_source || '',
      row.request_status || ''
    ];
  }) : [['No rows', '', '', '', '', '', '']];
}

function writeKeyValueBlock_(sheet, startRow, title, rows) {
  sheet.getRange(startRow, 1).setValue(title).setFontWeight('bold').setFontSize(12);
  const values = [['Metric', 'Value']].concat(rows);
  sheet.getRange(startRow + 1, 1, values.length, 2).setValues(values);
  styleHeaderRow_(sheet, startRow + 1, 2);
  return startRow + values.length + 3;
}

function writeTableBlock_(sheet, startRow, title, headers, rows) {
  sheet.getRange(startRow, 1).setValue(title).setFontWeight('bold').setFontSize(12);
  const dataRows = rows && rows.length ? rows : [headers.map(function(_, index) { return index === 0 ? 'No data yet' : ''; })];
  const values = [headers].concat(dataRows);
  sheet.getRange(startRow + 1, 1, values.length, headers.length).setValues(values);
  styleHeaderRow_(sheet, startRow + 1, headers.length);
  return startRow + values.length + 3;
}

function formatMoney_(value, currency) {
  return roundCurrency_(value) + ' ' + String(currency || getReportingCurrency_()).trim();
}

function formatPercent_(value) {
  return roundMetric_(Number(value || 0) * 100) + '%';
}

function roundMetric_(value) {
  return Math.round(Number(value || 0) * 100) / 100;
}

function matchesSearchTerm_(value, loweredTerm) {
  return String(value || '').toLowerCase().indexOf(loweredTerm) !== -1;
}

function getScriptTimeZone_() {
  return Session.getScriptTimeZone() || 'Europe/Tbilisi';
}

function parseIncomingPayload_(e) {
  if (!e) return {};

  if (e.postData && e.postData.contents) {
    const contentType = String((e.postData.type || '')).toLowerCase();

    if (contentType.indexOf('application/json') !== -1) {
      return JSON.parse(e.postData.contents);
    }

    // Fallback for x-www-form-urlencoded or text payloads.
    if (e.parameter && Object.keys(e.parameter).length) {
      return e.parameter;
    }
  }

  return e.parameter || {};
}

function escapeHtmlForTemplate_(value) {
  return String(value == null ? '' : value).replace(/[&<>"']/g, function(character) {
    return {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      '\'': '&#39;'
    }[character] || character;
  });
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

const PMS_EXPOSED_ADMIN_FUNCTION_NAMES = [
  'getAdminPanelBootstrapData',
  'getPmsDeferredBootstrapData',
  'getCommercialControlsData',
  'getRateBoardData',
  'getRateBoardCommercialSignals',
  'getRateCalendarData',
  'getBaseRatesData',
  'adminApplyRateLadder',
  'adminSaveRateCalendarCell',
  'adminQueueOtaFollowUpFromRateBoard',
  'getRecentBookingAlerts',
  'adminSaveBaseRate',
  'adminSaveBaseRatesBulk',
  'adminSaveCommercialControl',
  'adminSaveBulkDateRateOverrides',
  'adminSetCommercialControlActive',
  'getEventFlagsData',
  'adminSaveEventFlag',
  'getCompetitorTrackerData',
  'adminSaveCompetitorTrackerEntry',
  'getRecommendationActionLogData',
  'adminSaveRecommendationActionLog',
  'getOtaWorkflowData',
  'adminSaveOtaWorkflowStatus',
  'getPickupPaceData',
  'getDemandScoreData',
  'getPricingRecommendationData',
  'adminApplyManualCommercialShortcut',
  'adminCreateManualBooking',
  'adminGetBookingDetail',
  'adminGetRoomAssignmentOptions',
  'adminAssignBookingRoom',
  'adminUpdateBookingStatus',
  'adminMarkBookingPaid',
  'adminUpdateBooking',
  'adminCancelBooking',
  'adminCheckAvailability',
  'adminAddBlock',
  'adminSearchBookings',
  'getFrontDeskDashboardData',
  'getArrivalsPlanningData',
  'getDeparturesForDate',
  'getInHouseForDate',
  'getAvailabilityPlannerData',
  'getReportsDashboardData',
  'adminRefreshReporting',
  'adminRefreshAvailabilityCache',
  'adminRebuildDailyStats',
  'adminSetupOperationalSafetyTriggers',
  'adminVerifyWebsiteEmailProcessorTrigger',
  'getPublicDeploymentSafetyChecklist',
  'adminRunPmsHealthCheck',
  'adminSendDailyPmsSummaryEmail',
  'adminSendPendingWebsiteBookingEmails',
  'adminRetryFailedWebsiteBookingEmails',
  'adminGetWebsiteFeedbackDashboard'
];

PMS_EXPOSED_ADMIN_FUNCTION_NAMES.forEach(function(functionName) {
  const original = globalThis[functionName];
  if (typeof original !== 'function' || original.__rozaPmsAccessWrapped) return;
  const wrapped = function() {
    requirePmsAdminAccess_(arguments.length ? arguments[0] : {});
    return original.apply(this, arguments);
  };
  wrapped.__rozaPmsAccessWrapped = true;
  globalThis[functionName] = wrapped;
});
